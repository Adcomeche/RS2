VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSatellitareEvento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public SQL_GETLIST As String
Public SQL_FIND As String
Public SQL_UPDATE As String
Public SQL_INSERT As String
Public SQL_GET_TS As String
Private objLogger As CLogger
Private SQL_KNOSOS_GET_EVENTS As String
Private SQL_OESIA_GET_EVENTS As String
Private SQL_KNOSOS_DELETE_EVENT As String
Private SQL_KNOSOS_GET_KEY_CONSEGNA_PRODOTTO_RESOIL As String
Private SQL_KNOSOS_GET_KEY_VIAGGIO_RESOIL As String
Private SQL_KNOSOS_GET_LAST_EVENT_DATE As String
Private SQL_KNOSOS_GET_IMPORTO_PEDAGGIO As String
Private SQL_KNOSOS_GET_COUNT_SCARICHI_NONCONCLUSI As String
Private SQL_KNOSOS_GET_COUNT_SCARICHI_NONCONCLUSI_VVT As String
Private SQL_UPDATE_VIAGGIO_MODIFICATO As String
Private SQL_GET_SELECTION_EVENTS As String
Private SQL_SELECT_KM_EVENTS As String
Private SQL_GET_KM_EVENT As String

Private SQL_SMS_GET_EVENTS As String
Private SQL_GET_EVENTS_PALM As String
Private SQL_GET_EVENTOS_PALM_TOTAL As String
Private SQL_SMS_GET_LIST_AGG As String
Private SQL_PALM_GET_ORD As String
Private SQL_PALM_GET_CONSEGNA As String
Private SQL_PALM_GET_CUANTOSCLI As String
Private SQL_PALM_GET_ORDENESCLI As String
'--SM 15/02/2010
'--Construcción de la tabla GES
Private ADZ As CDestinazione
Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Private SQL_GES_001 As String
Private SQL_GES_007 As String
Private SQL_GES_campos As String
Private SQL_GES_values As String
Private SQL_GES_002003 As String
Private SQL_GES_CUCARGA As String
Private SQL_GES_005006 As String
Private SQL_GES_008 As String
Private SQL_GES_HorariosCli As String
Private Service As VLocalize.VLocalize
Private SQL_POS_CLI As String
Private SQL_POS_FAC As String


'-- DP 15/05/08
'-- Sostituisce SQL_SMS_GET_LIST_AGG
Private SQL_SMS_GET_LIST_AGG_2 As String

Private SQL_DEUSTO_GET_VIAGGI As String
Private SQL_SAT_INSERT_POS As String
Private SQL_SAT_GET_MAX_POS_TRASM As String
Private SQL_SAT_GET_POS_TRASM_LIST As String
Private SQL_SAT_UPDATE_POS_STATUS As String
Private SQL_SAT_GET_LAST_POS As String
Private SQL_SAT_GET_LAST_POS_2 As String
Private SQL_SAT_GET_LAST_POS_3 As String
Private SQL_SAT_GET_LAST_POS_4 As String

'-- DP 20/06/08
'-- Query per controllo SMS doppi
Private SQL_FIND_DUBLE_SMS As String

Private Const ALIQUOTA_IVA_PEDAGGIO As Single = 16
Private Const KNOSOS_VIAGGIO_STATO_TERMINATO = 9

Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1
Private mcolListFields As Collection

Private Const VIAGGIO_STATO_PENDENTE = 0
Private Const VIAGGIO_STATO_INIZIOVIAGGIO = 1
Private Const VIAGGIO_STATO_CARICO = 2
Private Const VIAGGIO_STATO_INCORSO = 3
Private Const VIAGGIO_STATO_CONCLUSO = 4
Private Const VIAGGIO_STATO_CANCELLATO = 5

Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)

Private Const SQL_CONSEGNA_IN_ORDINE = "SELECT 1 FROM VIAGVVC0 WHERE VVCIDVVT = ? AND VVCNRSEQ < ? AND " & _
                                    "           (VVCFLST5 = 'N' OR VVCFLST5 IS NULL) AND VVCANNUL = 0"

Private Const SQL_GET_SCARICHI_NON_CONCLUSI = "SELECT VIAGVVP0.VVPIDVVC, VIAGVVP0.VVPIDVVP, VIAGSEV0.SEVIDVVP " & _
                                            "  FROM VIAGVVP0 LEFT OUTER JOIN VIAGSEV0 ON " & _
                                            "       VIAGVVP0.VVPIDVVT = VIAGSEV0.SEVIDVVT AND " & _
                                            "       VIAGVVP0.VVPIDVVC = VIAGSEV0.SEVIDVVC AND " & _
                                            "       VIAGVVP0.VVPIDVVP = VIAGSEV0.SEVIDVVP " & _
                                            "  WHERE VVPIDVVT = ? AND VVPANNUL = 0"

Private Const SQL_GET_KM_VIAGGIO = "SELECT SUM SEVNRKM FROM VIAGSEV0 " & _
                                "   WHERE SEVIDVVT = ? AND SEVIDEVE = 'SI' AND SEVANNUL = 0"

Private Const SQL_GET_EVENTI_PRECEDENTI = "SELECT 1 FROM EventosMBFlotas WHERE CodigoEvento = ? AND CodigoServicio = ? AND IdRegistro < ?"

Private Const SQL_GET_VG_CARICATI_GIORNO_AUT = "SELECT VVTIDVVT FROM VIAGVVT0 WHERE VVTIDAUT = ? AND VVTDTINV = ? AND VVTFLST3 = 'S' AND VVTANNUL = 0 ORDER BY VVTHRINV"

' Restituisce True se non ci sono altri viaggi per una AUT in un giorno con VVTFLST3 = "S"
Private Function varGetViaggiCaricatiDelGiorno(varIDAUT As Variant, varDTINV As Variant) As Variant
    Dim varBuffer As Variant
    
    On Error GoTo ErrHandler
    
    If DS.DA.QueryCommand(SQL_GET_VG_CARICATI_GIORNO_AUT, varBuffer, 0, varIDAUT, varDTINV) Then
        varGetViaggiCaricatiDelGiorno = varBuffer
    Else
        GoTo ErrHandler
    End If
    
    Exit Function
    
ErrHandler:
    varGetViaggiCaricatiDelGiorno = Empty
End Function

Public Function SAT_Get_Max_Pos() As Variant
    On Error GoTo ErrHandler
    
    Dim varBuffer As Variant
    Dim lngHowMany As Long
    
    If Not DS.DA.QueryData(SQL_SAT_GET_MAX_POS_TRASM, varBuffer, lngHowMany) Then
        GoTo ErrHandler
    End If
    
    If IsEmpty(varBuffer) Then
        SAT_Get_Max_Pos = 0
    Else
        SAT_Get_Max_Pos = varBuffer(0, 0)
    End If
    
    Exit Function
    
ErrHandler:
    SAT_Get_Max_Pos = Empty
End Function

Public Function SAT_Get_Last_Pos(IDAutomezzo As Variant, Data As Variant, Ora As Variant, ByRef CoordX As Variant, ByRef CoordY As Variant) As Boolean
    On Error GoTo ErrHandler
    
    Dim varBuffer As Variant
    Dim lngHowMany As Long
    Dim i As Long
    Dim strDataOra As String
    
    If Not IsDate(Data) Then
        GoTo ErrHandler
    End If
    If Not IsDate(Ora) Then
        GoTo ErrHandler
    End If
    
    strDataOra = Format(Data, "yyyymmdd") & Format(Ora, "hhnnss")
    
    If Not DS.DA.QueryData(SQL_SAT_GET_LAST_POS & DS.DA.MakeSQL(IDAutomezzo, dtAlfa) & SQL_SAT_GET_LAST_POS_2 & DS.DA.MakeSQL(DateAdd("d", -2, Data), dtDate) & SQL_SAT_GET_LAST_POS_3 & DS.DA.MakeSQL(Data, dtDate) & SQL_SAT_GET_LAST_POS_4, varBuffer, lngHowMany) Then
        GoTo ErrHandler
    End If
    If IsEmpty(varBuffer) Then
        CoordX = Empty
        CoordY = Empty
    Else
        If Not IsDate(varBuffer(2, 0)) Then
            CoordX = Empty
            CoordY = Empty
        Else
            If Format(varBuffer(2, 0), "yyyymmdd") < Format(Data, "yyyymmdd") Then
            
                CoordX = varBuffer(0, 0)
                CoordY = varBuffer(1, 0)
            Else
                For i = 0 To UBound(varBuffer, 2)
                    If Format(varBuffer(3, i), "hhnnss") <= Format(Ora, "hhnnss") Then
                        CoordX = varBuffer(0, i)
                        CoordY = varBuffer(1, i)
                        Exit For
                    End If
                Next i
            End If
        End If
    End If
        
    SAT_Get_Last_Pos = True
    
    Exit Function
    
ErrHandler:
    CoordX = Empty
    CoordY = Empty
    SAT_Get_Last_Pos = False
End Function
Public Function SAT_Get_Pos_List(MaxID As Variant) As Variant
    On Error GoTo ErrHandler
    
    Dim varBuffer As Variant
    Dim lngHowMany As Long
    
    If Not DS.DA.QueryCommand(SQL_SAT_GET_POS_TRASM_LIST, varBuffer, lngHowMany, MaxID) Then
        GoTo ErrHandler
    End If
    
    If IsEmpty(varBuffer) Then
        GoTo ErrHandler
    End If
        
    SAT_Get_Pos_List = varBuffer
    
    Exit Function
    
ErrHandler:
    SAT_Get_Pos_List = Empty
End Function
Public Function SAT_Update_Pos_Status(MaxID As Variant) As Variant
    On Error GoTo ErrHandler
    
    Dim varBuffer As Variant
    Dim lngHowMany As Long
    
    If Not DS.DA.ExecuteCommand(SQL_SAT_UPDATE_POS_STATUS, MaxID) Then
        GoTo ErrHandler
    End If
    
    SAT_Update_Pos_Status = True
    
    Exit Function
    
ErrHandler:
    SAT_Update_Pos_Status = False
End Function

Public Function Find() As Boolean
    Find = DS.Find(Me.SQL_FIND)
End Function

Public Function GetList(Optional ActiveRecordsOnly As Boolean = False) As Boolean
    GetList = DS.GetList(SQL_GETLIST, ActiveRecordsOnly)
End Function
Public Function Insert() As Boolean
    Insert = DS.Insert(Me.SQL_INSERT)
End Function
Public Function BeginSave(ByRef blnConcurErr As Boolean) As Boolean
    Dim blnSaveOk As Boolean
    
    blnSaveOk = DS.BeginSave(SQL_UPDATE, SQL_GET_TS, blnConcurErr)
    
    BeginSave = True
    
End Function

Private Sub Class_Initialize()
    SQL_FIND = "SELECT SEVIDSEV, SEVIDVVT, SEVIDVVC, SEVIDVVP," & _
                " SEVIDEVE, SEVDTEVE, SEVQTLTC, SEVQTLTS, SEVNRKM," & _
                " SEVFLSAT, TIMESTMP FROM VIAGSEV0 WHERE "
    SQL_GETLIST = "SELECT SEVIDSEV, SEVIDVVT, SEVIDVVC, SEVIDVVP FROM VIAGSEV0 "
    SQL_UPDATE = "UPDATE VIAGSEV0 SET "
    SQL_INSERT = "INSERT INTO VIAGSEV0 "
    SQL_GET_TS = "SELECT TIMESTMP FROM VIAGSEV0 WHERE "
    
    SQL_SMS_GET_EVENTS = "SELECT SMSIDSMS, SMSNRTEL, SMSDTSMD, SMSDTLOC, " & _
                         "       SMSTESTO " & _
                         "  FROM VIAGSMS0 " & _
                         " WHERE SMSIDSMS > ? " & _
                         " ORDER BY SMSIDSMS "
    SQL_POS_CLI = "SELECT ISNULL(ADZCGPSX,0) As ADZCGPSX, ISNULL(ADZCGPSY,0)As ADZCGPSY From dbo.ANAGADZ0 WHERE  (ADZIDADZ = ?)"
    
    SQL_POS_FAC = "SELECT     dbo.MULTSDE0.SDECODSO, dbo.ANAGADE0.ADECGPSX, dbo.ANAGADE0.ADECGPSY FROM dbo.ANAGADE0 INNER JOIN " & _
                  "dbo.MULTSDE0 ON dbo.ANAGADE0.ADEIDADE = dbo.MULTSDE0.SDEIDADE WHERE (dbo.MULTSDE0.SDEIDASC = '901') AND (dbo.MULTSDE0.SDECODSO = ?)"
        
    SQL_GET_EVENTS_PALM = "SELECT     IdRegistro, IdTareaMB, IdUsuarioMB, CodigoVehiculo, CodigoEvento, CodigoTarea, CodigoServicio, Productos, FechaEvento, FechaServidor, " & _
                            "HoraDescarga , Metros, Latitud, Longitud, EntaEnLoc, Capturado, CONVERT(varchar, FechaEvento, 112) AS FechaEv, LEFT(CONVERT(varchar, FechaEvento, 24), 5) AS HoraEv  " & _
                            "From dbo.EventosMBFlotasManual " & _
                            "WHERE (IdRegistro = " & _
                            "(SELECT MIN(IdRegistro) " & _
                            "From EventosMBFlotasManual " & _
                            "  WHERE      Capturado = 0 AND ((CodigoEvento = '003') OR (CodigoEvento = '005'))))"
    
    SQL_GET_EVENTOS_PALM_TOTAL = "SELECT     IdRegistro, IdTareaMB, IdUsuarioMB, CodigoVehiculo, CodigoEvento, CodigoTarea, CodigoServicio, Productos, FechaEvento, FechaServidor, " & _
                            "HoraDescarga , Metros, Latitud, Longitud, EntaEnLoc, Capturado, CONVERT(varchar, FechaEvento, 112) AS FechaEv, LEFT(CONVERT(varchar, FechaEvento, 24), 5) AS HoraEv  " & _
                            "From dbo.EventosMBFlotasManual " & _
                            "WHERE (IdRegistro = " & _
                            "(SELECT MIN(IdRegistro) " & _
                            "From EventosMBFlotasManual " & _
                            "  WHERE      Capturado = 0 )) "
                            
    SQL_PALM_GET_ORD = "SELECT VVCIDVVC " & _
                       "FROM VIAGVVC0 " & _
                       "WHERE (VVCIDVVT = ?)"
                       
    SQL_PALM_GET_CONSEGNA = "SELECT VVCIDVVC From VIAGVVC0 WHERE (VVCIDOCT = ?) AND (VVCIDVVT = ?)"
    
    SQL_PALM_GET_CUANTOSCLI = "SELECT VVCIDADZ FROM dbo.VIAGVVC0 WHERE (VVCIDVVT = ?) AND (VVCIDOCT = ?)"
    
    SQL_PALM_GET_ORDENESCLI = "SELECT VVCIDOCT FROM dbo.VIAGVVC0 WHERE (VVCIDVVT = ?) AND (VVCIDADZ = ?) ORDER BY VVCFLST5"
    
    SQL_SMS_GET_LIST_AGG = "SELECT VVTIDAM1, VVCIDVVT, VVCIDVVC, VVTFLST3, " & _
                           "       VVCFLST5, VVTHRINC, CASE VVCHRPRE WHEN '01/01/1900' THEN VVCTEMVG ELSE VVCHRPRE END as ORA_PREV, VVCTEMVG  " & _
                           "  FROM VIAGVVT0 INNER JOIN VIAGVVC0 ON VVTIDVVT = VVCIDVVT " & _
                           " WHERE (VVTDTINV = ?) AND (VVTANNUL = 0) AND (VVCANNUL = 0) " & _
                           " ORDER BY VVTIDAM1, VVTDTINC, VVTHRINC, ORA_PREV "
                    
    
    SQL_GES_001 = " SELECT     TOP 1 PERCENT dbo.VIAGVVT0.VVTIDVVT, " & _
                    "                      CASE dbo.VIAGVVC0.VVCHRPRE WHEN '01/01/1900' THEN dbo.VIAGVVC0.VVCTEMVG ELSE dbo.VIAGVVC0.VVCHRPRE END AS HORAENT, " & _
                    "                      dbo.VIAGVVT0.VVTHRINC AS HoraCarga, dbo.ANAGAAU0.AAUNRCEL " & _
                    "FROM         dbo.VIAGVVT0 INNER JOIN " & _
                    "                      dbo.VIAGVVC0 ON dbo.VIAGVVT0.VVTIDVVT = dbo.VIAGVVC0.VVCIDVVT INNER JOIN " & _
                    "                      dbo.ANAGAAU0 ON dbo.VIAGVVT0.VVTIDAU1 = dbo.ANAGAAU0.AAUIDAAU " & _
                    "WHERE     (dbo.VIAGVVT0.VVTANNUL = 0) AND (dbo.VIAGVVC0.VVCANNUL = 0) AND (dbo.VIAGVVT0.VVTIDASC = '038') AND (dbo.VIAGVVC0.VVCFLST5 <> 'S' OR " & _
                    "                      dbo.VIAGVVC0.VVCFLST5 IS NULL)  AND (LEFT(dbo.VIAGVVT0.VVTDTRIF, 10) = LEFT(GETDATE(), 10)) AND " & _
                    "                      (dbo.ANAGAAU0.AAUNRCEL = ?) " & _
                    "ORDER BY dbo.VIAGVVT0.VVTHRINC "
                    
    
    SQL_GES_007 = "SELECT   TOP 1 PERCENT dbo.VIAGVVT0.VVTIDVVT AS VIAJE, dbo.VIAGVVC0.VVCIDADZ AS CLIENTE, " & _
                      "CASE dbo.VIAGVVC0.VVCHRPRE WHEN '01/01/1900' THEN dbo.VIAGVVC0.VVCTEMVG ELSE dbo.VIAGVVC0.VVCHRPRE END AS HORAENT, " & _
                      "dbo.VIAGVVT0.VVTHRINC AS HoraCarga, dbo.ANAGAAU0.AAUNRCEL " & _
                      "FROM         dbo.VIAGVVT0 INNER JOIN " & _
                      "dbo.VIAGVVC0 ON dbo.VIAGVVT0.VVTIDVVT = dbo.VIAGVVC0.VVCIDVVT INNER JOIN " & _
                      "dbo.ANAGAAU0 ON dbo.VIAGVVT0.VVTIDAU1 = dbo.ANAGAAU0.AAUIDAAU " & _
                      "WHERE     (dbo.VIAGVVT0.VVTANNUL = 0) AND (dbo.VIAGVVT0.VVTIDASC = '038')AND (dbo.VIAGVVC0.VVCANNUL = 0) AND (dbo.VIAGVVC0.VVCFLST5 = 'S') AND (dbo.VIAGVVT0.VVTDTRIF = ?) AND " & _
                      "(dbo.ANAGAAU0.AAUNRCEL = ?) " & _
                      "ORDER BY CASE dbo.VIAGVVC0.VVCHRPRE WHEN '01/01/1900' THEN dbo.VIAGVVC0.VVCTEMVG ELSE dbo.VIAGVVC0.VVCHRPRE END DESC "
                      
'  SQL_GES_002003 = " SELECT     dbo.VIAGVVT0.VVTIDAEE AS GESCODTR, dbo.ORDIOCT0.OCTRIFE1 + dbo.ORDIOCR0.OCRRIFE2 AS GESNUMPD, dbo.ORDIOCT0.OCTRAUT4 AS GESNUMPR, SUBSTRING(dbo.VIAGVVC0.VVCIDADZ, 3, 8) AS GESCODDE, LEFT(dbo.MULTSPR0.SPRCODSO, 7) " & _
'                      "AS GESPRCOD, dbo.MULTSDE0.SDECODSO AS GESFACID, dbo.ORDIOCR0.OCRQCOLT AS GESCANPR, dbo.ORDIOCR0.OCRIDAUM AS GESUNMED, CONVERT(varchar, dbo.VIAGVVC0.VVCDTCON, 112) AS GESDYDES, dbo.VIAGVVP0.VVPQTRLT AS GESCARTR, dbo.VIAGVVP0.VVPQCOLT AS GESDESTR, " & _
'                      "CONVERT(varchar, dbo.VIAGVVT0.VVTDTINC, 112) AS GESDYTR0, LEFT(CONVERT(varchar, dbo.VIAGVVT0.VVTHRINC, 24), 5) AS GESHRTR0, dbo.MULTSUT0.SUTCODSO AS GESVHIR0, dbo.VIAGVVT0.VVTSEQVI AS GESVIAGR, 'ES' + dbo.ANAGAAU0.AAUNRCID AS GESCONR0, " & _
'                      "dbo.ANAGDED0.DEDKMATT AS GESKMCEP, LEFT(dbo.VIAGVVC0.VVCIDVVC, 2) AS GESNUMVI, dbo.VIAGVVC0.VVCNOTE1 AS GESNOTES, dbo.VIAGVVT0.VVTKMDEF, CASE dbo.viagvvc0.vvcdtpre WHEN '01/01/1900' THEN CONVERT(varchar, dbo.VIAGVVC0.VVCDTARR, 112) " & _
'                      "ELSE CONVERT(varchar, dbo.VIAGVVC0.VVCDTPRE, 112) END AS GESDYACT, CASE dbo.viagvvc0.vvchrpre WHEN '01/01/1900' THEN LEFT(CONVERT(varchar, dbo.VIAGVVC0.VVCHRARR, 24), 5) ELSE LEFT(CONVERT(varchar," & _
'                      "dbo.VIAGVVC0.VVCHRPRE, 24), 5) END AS GESHRACT, CASE viagvvc0.vvcflst5 WHEN 'S' THEN LEFT(CONVERT(varchar, dbo.VIAGVVC0.VVCHRARR, 24), 5) ELSE '' END AS GESHR00, CASE viagvvc0.vvcflst5 WHEN 'S' THEN CONVERT(varchar, dbo.VIAGVVC0.VVCDTARR, 112) " & _
'                      "Else '' END AS GESDYR00 FROM  dbo.ANAGDED0 RIGHT OUTER JOIN dbo.VIAGVVT0 INNER JOIN  dbo.VIAGVVC0 ON dbo.VIAGVVT0.VVTIDVVT = dbo.VIAGVVC0.VVCIDVVT INNER JOIN  dbo.VIAGVVP0 ON dbo.VIAGVVC0.VVCIDVVT = dbo.VIAGVVP0.VVPIDVVT AND dbo.VIAGVVC0.VVCIDVVC = dbo.VIAGVVP0.VVPIDVVC INNER JOIN " & _
'                      "dbo.ANAGAAM0 ON dbo.VIAGVVT0.VVTIDAM1 = dbo.ANAGAAM0.AAMIDAAM INNER JOIN dbo.ORDIOCT0 ON dbo.VIAGVVC0.VVCIDOCT = dbo.ORDIOCT0.OCTIDOCT INNER JOIN dbo.ORDIOCR0 ON dbo.VIAGVVP0.VVPIDOCT = dbo.ORDIOCR0.OCRIDOCT AND dbo.VIAGVVP0.VVPIDOCR = dbo.ORDIOCR0.OCRIDOCR AND " & _
'                      "dbo.ORDIOCT0.OCTIDOCT = dbo.ORDIOCR0.OCRIDOCT INNER JOIN dbo.ANAGADZ0 ON dbo.VIAGVVC0.VVCIDADZ = dbo.ANAGADZ0.ADZIDADZ LEFT OUTER JOIN dbo.ANAGAAU0 ON dbo.VIAGVVT0.VVTIDAU1 = dbo.ANAGAAU0.AAUIDAAU LEFT OUTER JOIN " & _
'                      "dbo.MULTSDE0 ON dbo.VIAGVVT0.VVTIDADE = dbo.MULTSDE0.SDEIDADE LEFT OUTER JOIN dbo.ANAGAUT0 LEFT OUTER JOIN dbo.MULTSUT0 RIGHT OUTER JOIN dbo.ANAGADE0 ON dbo.MULTSUT0.SUTIDASC = dbo.ANAGADE0.ADEIDASC ON dbo.ANAGAUT0.AUTIDAUT = dbo.MULTSUT0.SUTIDAUT ON " & _
'                      "dbo.MULTSPR0 ON dbo.ORDIOCR0.OCRIDAPR = dbo.MULTSPR0.SPRIDAPR WHERE     (dbo.VIAGVVC0.VVCANNUL = 0) AND (dbo.VIAGVVP0.VVPANNUL = 0) AND (dbo.MULTSPR0.SPRIDASC = '901') AND (dbo.MULTSDE0.SDEIDASC = '901') AND (dbo.VIAGVVT0.VVTIDVVT = ?) "
'
                      
                      
                      

  SQL_GES_002003 = "SELECT     dbo.VIAGVVT0.VVTIDAEE AS GESCODTR, dbo.ORDIOCT0.OCTRIFE1 + dbo.ORDIOCR0.OCRRIFE2 AS GESNUMPD," & _
                      "dbo.ORDIOCT0.OCTRAUT4 AS GESNUMPR, SUBSTRING(dbo.VIAGVVC0.VVCIDADZ, 3, 8) AS GESCODDE,  " & _
                      "left(dbo.MULTSPR0.SPRCODSO,7) AS GESPRCOD, dbo.MULTSDE0.SDECODSO AS GESFACID," & _
                      "dbo.ORDIOCR0.OCRQCOLT AS GESCANPR, dbo.ORDIOCR0.OCRIDAUM AS GESUNMED, CONVERT(varchar, dbo.VIAGVVC0.VVCDTCON, 112) " & _
                      "AS GESDYDES, dbo.VIAGVVP0.VVPQTRLT AS GESCARTR, dbo.VIAGVVP0.VVPQCOLT AS GESDESTR, CONVERT(varchar, dbo.VIAGVVT0.VVTDTINC, 112) " & _
                      "AS GESDYTR0, LEFT(CONVERT(varchar, dbo.VIAGVVT0.VVTHRINC, 24), 5) AS GESHRTR0,  dbo.MULTSUT0.SUTCODSO AS GESVHIR0, " & _
                      "dbo.VIAGVVT0.VVTSEQVI AS GESVIAGR,  dbo.ANAGAAU0.AAUNRCID AS GESCONR0, dbo.ANAGDED0.DEDKMATT AS GESKMCEP, " & _
                      "LEFT(dbo.VIAGVVC0.VVCIDVVC, 2)AS GESNUMVI, VIAGVVC0.VVCNOTE1 AS GESNOTES, dbo.VIAGVVC0.VVCKMPAD As GESKMCR0, " & _
                      " CASE dbo.viagvvc0.vvcdtpre WHEN '01/01/1900' THEN CONVERT(varchar, dbo.VIAGVVC0.VVCDTARR, 112) ELSE CONVERT(varchar, dbo.VIAGVVC0.VVCDTPRE, 112) END AS GESDYACT,CASE dbo.viagvvc0.vvchrpre WHEN '01/01/1900' THEN LEFT(CONVERT(varchar, dbo.VIAGVVC0.VVCHRARR, 24), 5) ELSE LEFT(CONVERT(varchar, dbo.VIAGVVC0.VVCHRPRE, 24), 5) END AS GESHRACT " & _
                      "  , CASE viagvvc0.vvcflst5 WHEN 'S' THEN LEFT(CONVERT(varchar, dbo.VIAGVVC0.VVCHRARR, 24), 5) ELSE '' END AS GESHR00, CASE viagvvc0.vvcflst5 WHEN 'S' THEN CONVERT(varchar, dbo.VIAGVVC0.VVCDTARR, 112) Else '' END AS GESDYR00, dbo.VIAGVVT0.VVTFLUDG      FROM         dbo.ANAGDED0 RIGHT OUTER JOIN " & _
                      " dbo.VIAGVVT0 INNER JOIN " & _
                      " dbo.VIAGVVC0 ON dbo.VIAGVVT0.VVTIDVVT = dbo.VIAGVVC0.VVCIDVVT INNER JOIN dbo.VIAGVVP0 ON dbo.VIAGVVC0.VVCIDVVT = dbo.VIAGVVP0.VVPIDVVT AND dbo.VIAGVVC0.VVCIDVVC = dbo.VIAGVVP0.VVPIDVVC INNER JOIN " & _
                      " dbo.ANAGAAM0 ON dbo.VIAGVVT0.VVTIDAM1 = dbo.ANAGAAM0.AAMIDAAM INNER JOIN dbo.ORDIOCT0 ON dbo.VIAGVVC0.VVCIDOCT = dbo.ORDIOCT0.OCTIDOCT INNER JOIN " & _
                      " dbo.ORDIOCR0 ON dbo.VIAGVVP0.VVPIDOCT = dbo.ORDIOCR0.OCRIDOCT AND dbo.VIAGVVP0.VVPIDOCR = dbo.ORDIOCR0.OCRIDOCR AND dbo.ORDIOCT0.OCTIDOCT = dbo.ORDIOCR0.OCRIDOCT INNER JOIN " & _
                      " dbo.ANAGADZ0 ON dbo.VIAGVVC0.VVCIDADZ = dbo.ANAGADZ0.ADZIDADZ LEFT OUTER JOIN dbo.ANAGAAU0 ON dbo.VIAGVVT0.VVTIDAU1 = dbo.ANAGAAU0.AAUIDAAU LEFT OUTER JOIN " & _
                      " dbo.MULTSDE0 ON dbo.VIAGVVT0.VVTIDADE = dbo.MULTSDE0.SDEIDADE LEFT OUTER JOIN dbo.ANAGAUT0 LEFT OUTER JOIN dbo.MULTSUT0 RIGHT OUTER JOIN " & _
                      " dbo.ANAGADE0 ON dbo.MULTSUT0.SUTIDASC = dbo.ANAGADE0.ADEIDASC ON dbo.ANAGAUT0.AUTIDAUT = dbo.MULTSUT0.SUTIDAUT ON " & _
                      " dbo.VIAGVVT0.VVTIDADE = dbo.ANAGADE0.ADEIDADE AND dbo.VIAGVVT0.VVTIDAUT = dbo.ANAGAUT0.AUTIDAUT ON " & _
                      " dbo.ANAGDED0.DEDIDADE = dbo.VIAGVVT0.VVTIDADE AND dbo.ANAGDED0.DEDIDADZ = dbo.VIAGVVC0.VVCIDADZ LEFT OUTER JOIN " & _
                      " dbo.MULTSPR0 ON dbo.ORDIOCR0.OCRIDAPR = dbo.MULTSPR0.SPRIDAPR " & _
                      " WHERE     (dbo.VIAGVVC0.VVCANNUL = 0) AND (dbo.VIAGVVP0.VVPANNUL = 0) AND (dbo.MULTSPR0.SPRIDASC = '901') AND (dbo.MULTSDE0.SDEIDASC = '901') AND (dbo.VIAGVVT0.VVTIDVVT = ?) "


                               
                 
    SQL_GES_CUCARGA = "SELECT SUM(dbo.ANAGAPR0.APRDECON * dbo.ORDIOCR0.OCRQCOLT) AS PESOPR, " & _
                      "                      ANAGAAM0_1.AAMPORRI - dbo.ANAGAAM0.AAMTARAP - ANAGAAM0_1.AAMTARAP AS CMA " & _
                      "FROM         dbo.VIAGVVT0 INNER JOIN " & _
                      " dbo.VIAGVVC0 ON dbo.VIAGVVT0.VVTIDVVT = dbo.VIAGVVC0.VVCIDVVT INNER JOIN " & _
                      " dbo.VIAGVVP0 ON dbo.VIAGVVC0.VVCIDVVT = dbo.VIAGVVP0.VVPIDVVT AND dbo.VIAGVVC0.VVCIDVVC = dbo.VIAGVVP0.VVPIDVVC INNER JOIN " & _
                      " dbo.ORDIOCT0 ON dbo.VIAGVVC0.VVCIDOCT = dbo.ORDIOCT0.OCTIDOCT INNER JOIN " & _
                      " dbo.ORDIOCR0 ON dbo.VIAGVVP0.VVPIDOCT = dbo.ORDIOCR0.OCRIDOCT AND dbo.VIAGVVP0.VVPIDOCR = dbo.ORDIOCR0.OCRIDOCR AND  " & _
                      " dbo.ORDIOCT0.OCTIDOCT = dbo.ORDIOCR0.OCRIDOCT INNER JOIN " & _
                      " dbo.ANAGAPR0 ON dbo.ORDIOCR0.OCRIDAPR = dbo.ANAGAPR0.APRIDAPR INNER JOIN " & _
                      " dbo.ANAGAUT0 ON dbo.VIAGVVT0.VVTIDAUT = dbo.ANAGAUT0.AUTIDAUT INNER JOIN " & _
                      " dbo.ANAGAAM0 ON dbo.ANAGAUT0.AUTIDAM1 = dbo.ANAGAAM0.AAMIDAAM INNER JOIN " & _
                      " dbo.ANAGAAM0 ANAGAAM0_1 ON dbo.ANAGAUT0.AUTIDAM2 = ANAGAAM0_1.AAMIDAAM " & _
                      " GROUP BY dbo.VIAGVVT0.VVTIDVVT, dbo.ANAGAUT0.AUTPORKG - dbo.ANAGAAM0.AAMTARAP - ANAGAAM0_1.AAMTARAP, " & _
                      " ANAGAAM0_1.AAMPORRI - dbo.ANAGAAM0.AAMTARAP - ANAGAAM0_1.AAMTARAP " & _
                       "HAVING      (dbo.VIAGVVT0.VVTIDVVT = ? )"
                      
                   
'  SQL_GES_005006 = " SELECT     dbo.VIAGVVT0.VVTIDAEE AS GESCODTR, dbo.ORDIOCT0.OCTRIFE1 + dbo.ORDIOCR0.OCRRIFE2 AS GESNUMPD, dbo.ORDIOCT0.OCTRAUT4 AS GESNUMPR, SUBSTRING(dbo.VIAGVVC0.VVCIDADZ, 3, 8) AS GESCODDE, LEFT(dbo.MULTSPR0.SPRCODSO, 7) " & _
'                      "AS GESPRCOD, dbo.MULTSDE0.SDECODSO AS GESFACID, dbo.ORDIOCR0.OCRQCOLT AS GESCANPR, dbo.ORDIOCR0.OCRIDAUM AS GESUNMED, CONVERT(varchar, dbo.VIAGVVC0.VVCDTCON, 112) AS GESDYDES, dbo.VIAGVVP0.VVPQTRLT AS GESCARTR, dbo.VIAGVVP0.VVPQCOLT AS GESDESTR, " & _
'                      "CONVERT(varchar, dbo.VIAGVVT0.VVTDTINC, 112) AS GESDYTR0, LEFT(CONVERT(varchar, dbo.VIAGVVT0.VVTHRINC, 24), 5) AS GESHRTR0, dbo.MULTSUT0.SUTCODSO AS GESVHIR0, dbo.VIAGVVT0.VVTSEQVI AS GESVIAGR, 'ES' + dbo.ANAGAAU0.AAUNRCID AS GESCONR0, " & _
'                      "dbo.ANAGDED0.DEDKMATT AS GESKMCEP, LEFT(dbo.VIAGVVC0.VVCIDVVC, 2) AS GESNUMVI, dbo.VIAGVVC0.VVCNOTE1 AS GESNOTES, dbo.VIAGVVT0.VVTKMDEF, CASE dbo.viagvvc0.vvcdtpre WHEN '01/01/1900' THEN CONVERT(varchar, dbo.VIAGVVC0.VVCDTARR, 112) " & _
'                      "ELSE CONVERT(varchar, dbo.VIAGVVC0.VVCDTPRE, 112) END AS GESDYACT, CASE dbo.viagvvc0.vvchrpre WHEN '01/01/1900' THEN LEFT(CONVERT(varchar, dbo.VIAGVVC0.VVCHRARR, 24), 5) ELSE LEFT(CONVERT(varchar," & _
'                      "dbo.VIAGVVC0.VVCHRPRE, 24), 5) END AS GESHRACT, CASE viagvvc0.vvcflst5 WHEN 'S' THEN LEFT(CONVERT(varchar, dbo.VIAGVVC0.VVCHRARR, 24), 5) ELSE '' END AS GESHR00, CASE viagvvc0.vvcflst5 WHEN 'S' THEN CONVERT(varchar, dbo.VIAGVVC0.VVCDTARR, 112) Else '' END AS GESDYR00 FROM  dbo.ANAGDED0 RIGHT OUTER JOIN dbo.VIAGVVT0 INNER JOIN  dbo.VIAGVVC0 ON dbo.VIAGVVT0.VVTIDVVT = dbo.VIAGVVC0.VVCIDVVT INNER JOIN  dbo.VIAGVVP0 ON dbo.VIAGVVC0.VVCIDVVT = dbo.VIAGVVP0.VVPIDVVT AND dbo.VIAGVVC0.VVCIDVVC = dbo.VIAGVVP0.VVPIDVVC INNER JOIN " & _
'                      "dbo.ANAGAAM0 ON dbo.VIAGVVT0.VVTIDAM1 = dbo.ANAGAAM0.AAMIDAAM INNER JOIN dbo.ORDIOCT0 ON dbo.VIAGVVC0.VVCIDOCT = dbo.ORDIOCT0.OCTIDOCT INNER JOIN dbo.ORDIOCR0 ON dbo.VIAGVVP0.VVPIDOCT = dbo.ORDIOCR0.OCRIDOCT AND dbo.VIAGVVP0.VVPIDOCR = dbo.ORDIOCR0.OCRIDOCR AND " & _
'                      "dbo.ORDIOCT0.OCTIDOCT = dbo.ORDIOCR0.OCRIDOCT INNER JOIN dbo.ANAGADZ0 ON dbo.VIAGVVC0.VVCIDADZ = dbo.ANAGADZ0.ADZIDADZ LEFT OUTER JOIN dbo.ANAGAAU0 ON dbo.VIAGVVT0.VVTIDAU1 = dbo.ANAGAAU0.AAUIDAAU LEFT OUTER JOIN " & _
'                      "dbo.MULTSDE0 ON dbo.VIAGVVT0.VVTIDADE = dbo.MULTSDE0.SDEIDADE LEFT OUTER JOIN dbo.ANAGAUT0 LEFT OUTER JOIN dbo.MULTSUT0 RIGHT OUTER JOIN dbo.ANAGADE0 ON dbo.MULTSUT0.SUTIDASC = dbo.ANAGADE0.ADEIDASC ON dbo.ANAGAUT0.AUTIDAUT = dbo.MULTSUT0.SUTIDAUT ON " & _
'                      "dbo.MULTSPR0 ON dbo.ORDIOCR0.OCRIDAPR = dbo.MULTSPR0.SPRIDAPR WHERE     (dbo.VIAGVVC0.VVCANNUL = 0) AND (dbo.VIAGVVP0.VVPANNUL = 0) AND (dbo.MULTSPR0.SPRIDASC = '901') AND (dbo.MULTSDE0.SDEIDASC = '901') AND (dbo.VIAGVVT0.VVTIDVVT = ?) AND (dbo.ANAGADZ0.ADZIDADZ = ?) "
'
'
'

SQL_GES_005006 = "SELECT     dbo.VIAGVVT0.VVTIDAEE AS GESCODTR, dbo.ORDIOCT0.OCTRIFE1 + dbo.ORDIOCR0.OCRRIFE2 AS GESNUMPD," & _
                      " dbo.ORDIOCT0.OCTRAUT4 AS GESNUMPR, SUBSTRING(dbo.VIAGVVC0.VVCIDADZ, 3, 8) AS GESCODDE,  " & _
                      "  left(dbo.MULTSPR0.SPRCODSO,7) AS GESPRCOD, dbo.MULTSDE0.SDECODSO AS GESFACID," & _
                      " dbo.ORDIOCR0.OCRQCOLT AS GESCANPR, dbo.ORDIOCR0.OCRIDAUM AS GESUNMED, CONVERT(varchar, dbo.VIAGVVC0.VVCDTCON, 112) " & _
                      " AS GESDYDES, dbo.VIAGVVP0.VVPQTRLT AS GESCARTR, dbo.VIAGVVP0.VVPQCOLT AS GESDESTR, CONVERT(varchar, dbo.VIAGVVT0.VVTDTINC, 112) " & _
                      " AS GESDYTR0, LEFT(CONVERT(varchar, dbo.VIAGVVT0.VVTHRINC, 24), 5) AS GESHRTR0, dbo.MULTSUT0.SUTCODSO AS GESVHIR0, " & _
                      " dbo.VIAGVVT0.VVTSEQVI AS GESVIAGR, dbo.ANAGAAU0.AAUNRCID AS GESCONR0, dbo.ANAGDED0.DEDKMATT AS GESKMCEP, " & _
                      " LEFT(dbo.VIAGVVC0.VVCIDVVC, 2)AS GESNUMVI, VIAGVVC0.VVCNOTE1 AS GESNOTES, dbo.VIAGVVC0.VVCKMPAD, " & _
                      " CASE dbo.viagvvc0.vvcdtpre WHEN '01/01/1900' THEN CONVERT(varchar, dbo.VIAGVVC0.VVCDTARR, 112) ELSE CONVERT(varchar, dbo.VIAGVVC0.VVCDTPRE, 112) END AS GESDYACT, CASE dbo.viagvvc0.vvchrpre WHEN '01/01/1900' THEN LEFT(CONVERT(varchar, dbo.VIAGVVC0.VVCHRARR, 24), 5) ELSE LEFT(CONVERT(varchar, dbo.VIAGVVC0.VVCHRPRE, 24), 5) END AS GESHRACT " & _
                      " , CASE viagvvc0.vvcflst5 WHEN 'S' THEN LEFT(CONVERT(varchar, dbo.VIAGVVC0.VVCHRARR, 24), 5) ELSE '' END AS GESHR00, CASE viagvvc0.vvcflst5 WHEN 'S' THEN CONVERT(varchar, dbo.VIAGVVC0.VVCDTARR, 112) Else '' END AS GESDYR00 , dbo.VIAGVVT0.VVTFLUDG " & _
                    " FROM         dbo.ANAGDED0 RIGHT OUTER JOIN " & _
                    " dbo.VIAGVVT0 INNER JOIN " & _
                    " dbo.VIAGVVC0 ON dbo.VIAGVVT0.VVTIDVVT = dbo.VIAGVVC0.VVCIDVVT INNER JOIN dbo.VIAGVVP0 ON dbo.VIAGVVC0.VVCIDVVT = dbo.VIAGVVP0.VVPIDVVT AND dbo.VIAGVVC0.VVCIDVVC = dbo.VIAGVVP0.VVPIDVVC INNER JOIN " & _
                    " dbo.ANAGAAM0 ON dbo.VIAGVVT0.VVTIDAM1 = dbo.ANAGAAM0.AAMIDAAM INNER JOIN dbo.ORDIOCT0 ON dbo.VIAGVVC0.VVCIDOCT = dbo.ORDIOCT0.OCTIDOCT INNER JOIN " & _
                    " dbo.ORDIOCR0 ON dbo.VIAGVVP0.VVPIDOCT = dbo.ORDIOCR0.OCRIDOCT AND dbo.VIAGVVP0.VVPIDOCR = dbo.ORDIOCR0.OCRIDOCR AND dbo.ORDIOCT0.OCTIDOCT = dbo.ORDIOCR0.OCRIDOCT INNER JOIN " & _
                    " dbo.ANAGADZ0 ON dbo.VIAGVVC0.VVCIDADZ = dbo.ANAGADZ0.ADZIDADZ LEFT OUTER JOIN dbo.ANAGAAU0 ON dbo.VIAGVVT0.VVTIDAU1 = dbo.ANAGAAU0.AAUIDAAU LEFT OUTER JOIN " & _
                    " dbo.MULTSDE0 ON dbo.VIAGVVT0.VVTIDADE = dbo.MULTSDE0.SDEIDADE LEFT OUTER JOIN dbo.ANAGAUT0 LEFT OUTER JOIN dbo.MULTSUT0 RIGHT OUTER JOIN " & _
                    " dbo.ANAGADE0 ON dbo.MULTSUT0.SUTIDASC = dbo.ANAGADE0.ADEIDASC ON dbo.ANAGAUT0.AUTIDAUT = dbo.MULTSUT0.SUTIDAUT ON " & _
                    " dbo.VIAGVVT0.VVTIDADE = dbo.ANAGADE0.ADEIDADE AND dbo.VIAGVVT0.VVTIDAUT = dbo.ANAGAUT0.AUTIDAUT ON " & _
                    " dbo.ANAGDED0.DEDIDADE = dbo.VIAGVVT0.VVTIDADE AND dbo.ANAGDED0.DEDIDADZ = dbo.VIAGVVC0.VVCIDADZ LEFT OUTER JOIN " & _
                    " dbo.MULTSPR0 ON dbo.ORDIOCR0.OCRIDAPR = dbo.MULTSPR0.SPRIDAPR " & _
                    " WHERE     (dbo.VIAGVVC0.VVCANNUL = 0) AND (dbo.VIAGVVP0.VVPANNUL = 0) AND (dbo.MULTSPR0.SPRIDASC = '901') AND (dbo.MULTSDE0.SDEIDASC = '901') AND (dbo.VIAGVVT0.VVTIDVVT = ?)AND (dbo.ANAGADZ0.ADZIDADZ = ?)"

SQL_GES_008 = "SELECT     dbo.VIAGVVT0.VVTIDAEE AS GESCODTR, dbo.ORDIOCT0.OCTRIFE1 + dbo.ORDIOCR0.OCRRIFE2 AS GESNUMPD," & _
                      " dbo.ORDIOCT0.OCTRAUT4 AS GESNUMPR, SUBSTRING(dbo.VIAGVVC0.VVCIDADZ, 3, 8) AS GESCODDE,  " & _
                      "  left(dbo.MULTSPR0.SPRCODSO,7) AS GESPRCOD, dbo.MULTSDE0.SDECODSO AS GESFACID," & _
                      " dbo.ORDIOCR0.OCRQCOLT AS GESCANPR, dbo.ORDIOCR0.OCRIDAUM AS GESUNMED, CONVERT(varchar, dbo.VIAGVVC0.VVCDTCON, 112) " & _
                      " AS GESDYDES, dbo.VIAGVVP0.VVPQTRLT AS GESCARTR, dbo.VIAGVVP0.VVPQCOLT AS GESDESTR, CONVERT(varchar, dbo.VIAGVVT0.VVTDTINC, 112) " & _
                      " AS GESDYTR0, LEFT(CONVERT(varchar, dbo.VIAGVVT0.VVTHRINC, 24), 5) AS GESHRTR0, dbo.MULTSUT0.SUTCODSO AS GESVHIR0, " & _
                      " dbo.VIAGVVT0.VVTSEQVI AS GESVIAGR, dbo.ANAGAAU0.AAUNRCID AS GESCONR0, dbo.ANAGDED0.DEDKMATT AS GESKMCEP, " & _
                      " LEFT(dbo.VIAGVVC0.VVCIDVVC, 2)AS GESNUMVI, VIAGVVC0.VVCNOTE1 AS GESNOTES, dbo.VIAGVVT0.VVTKMDEF, " & _
                      " CASE dbo.viagvvc0.vvcdtpre WHEN '01/01/1900' THEN CONVERT(varchar, dbo.VIAGVVC0.VVCDTARR, 112) ELSE CONVERT(varchar, dbo.VIAGVVC0.VVCDTPRE, 112) END AS GESDYACT, CASE dbo.viagvvc0.vvchrpre WHEN '01/01/1900' THEN LEFT(CONVERT(varchar, dbo.VIAGVVC0.VVCHRARR, 24), 5) ELSE LEFT(CONVERT(varchar, dbo.VIAGVVC0.VVCHRPRE, 24), 5) END AS GESHRACT " & _
                      " , CASE viagvvc0.vvcflst5 WHEN 'S' THEN LEFT(CONVERT(varchar, dbo.VIAGVVC0.VVCHRARR, 24), 5) ELSE '' END AS GESHR00, CASE viagvvc0.vvcflst5 WHEN 'S' THEN CONVERT(varchar, dbo.VIAGVVC0.VVCDTARR, 112) Else '' END AS GESDYR00 " & _
                    " FROM         dbo.ANAGDED0 RIGHT OUTER JOIN " & _
                    " dbo.VIAGVVT0 INNER JOIN " & _
                    " dbo.VIAGVVC0 ON dbo.VIAGVVT0.VVTIDVVT = dbo.VIAGVVC0.VVCIDVVT INNER JOIN dbo.VIAGVVP0 ON dbo.VIAGVVC0.VVCIDVVT = dbo.VIAGVVP0.VVPIDVVT AND dbo.VIAGVVC0.VVCIDVVC = dbo.VIAGVVP0.VVPIDVVC INNER JOIN " & _
                    " dbo.ANAGAAM0 ON dbo.VIAGVVT0.VVTIDAM1 = dbo.ANAGAAM0.AAMIDAAM INNER JOIN dbo.ORDIOCT0 ON dbo.VIAGVVC0.VVCIDOCT = dbo.ORDIOCT0.OCTIDOCT INNER JOIN " & _
                    " dbo.ORDIOCR0 ON dbo.VIAGVVP0.VVPIDOCT = dbo.ORDIOCR0.OCRIDOCT AND dbo.VIAGVVP0.VVPIDOCR = dbo.ORDIOCR0.OCRIDOCR AND dbo.ORDIOCT0.OCTIDOCT = dbo.ORDIOCR0.OCRIDOCT INNER JOIN " & _
                    " dbo.ANAGADZ0 ON dbo.VIAGVVC0.VVCIDADZ = dbo.ANAGADZ0.ADZIDADZ LEFT OUTER JOIN dbo.ANAGAAU0 ON dbo.VIAGVVT0.VVTIDAU1 = dbo.ANAGAAU0.AAUIDAAU LEFT OUTER JOIN " & _
                    " dbo.MULTSDE0 ON dbo.VIAGVVT0.VVTIDADE = dbo.MULTSDE0.SDEIDADE LEFT OUTER JOIN dbo.ANAGAUT0 LEFT OUTER JOIN dbo.MULTSUT0 RIGHT OUTER JOIN " & _
                    " dbo.ANAGADE0 ON dbo.MULTSUT0.SUTIDASC = dbo.ANAGADE0.ADEIDASC ON dbo.ANAGAUT0.AUTIDAUT = dbo.MULTSUT0.SUTIDAUT ON " & _
                    " dbo.VIAGVVT0.VVTIDADE = dbo.ANAGADE0.ADEIDADE AND dbo.VIAGVVT0.VVTIDAUT = dbo.ANAGAUT0.AUTIDAUT ON " & _
                    " dbo.ANAGDED0.DEDIDADE = dbo.VIAGVVT0.VVTIDADE AND dbo.ANAGDED0.DEDIDADZ = dbo.VIAGVVC0.VVCIDADZ LEFT OUTER JOIN " & _
                    " dbo.MULTSPR0 ON dbo.ORDIOCR0.OCRIDAPR = dbo.MULTSPR0.SPRIDAPR " & _
                    " WHERE     (dbo.VIAGVVC0.VVCANNUL = 0) AND (dbo.VIAGVVP0.VVPANNUL = 0) AND (dbo.MULTSPR0.SPRIDASC = '901') AND (dbo.MULTSDE0.SDEIDASC = '901') AND (dbo.VIAGVVT0.VVTIDVVT = ?)AND (dbo.VIAGVVP0.VVPIDOCT = ?)"


'MIRA LOS HORARIOS DE APERTURA DE LOS CLIENTES
'  SQL_GES_HorariosCli = "SELECT     dbo.ORDIOCT0.OCTHRFR1 AS Desde1 ,  dbo.ORDIOCT0.OCTHRTO1 AS Hasta1 , " & _
'                        "dbo.ORDIOCT0.OCTHRFR2     AS  Desde2 ,  dbo.ORDIOCT0.OCTHRTO2  AS Hasta2, dbo.ORDIOCT0.OCTHRFR3 AS Desde3 ,  dbo.ORDIOCT0.OCTHRTO3 As Hasta3 " & _
'                        "FROM         dbo.VIAGVVT0 INNER JOIN " & _
'                        "dbo.VIAGVVC0 ON dbo.VIAGVVT0.VVTIDVVT = dbo.VIAGVVC0.VVCIDVVT INNER JOIN " & _
'                        "dbo.VIAGVVP0 ON dbo.VIAGVVC0.VVCIDVVT = dbo.VIAGVVP0.VVPIDVVT AND dbo.VIAGVVC0.VVCIDVVC = dbo.VIAGVVP0.VVPIDVVC INNER JOIN " & _
'                        "dbo.ORDIOCT0 ON dbo.VIAGVVC0.VVCIDOCT = dbo.ORDIOCT0.OCTIDOCT " & _
'                        "WHERE     (dbo.VIAGVVT0.VVTIDVVT = ?) AND (dbo.VIAGVVC0.VVCIDADZ = ?)"
'MIRA LOS HORARIOS DE SOLCITUD DE EL PRODUCTO.

SQL_GES_HorariosCli = " SELECT     dbo.ORDIOCT0.OCTHRFR1, dbo.ORDIOCT0.OCTHRTO1, dbo.ORDIOCT0.OCTHRFR2, dbo.ORDIOCT0.OCTHRTO2, dbo.ORDIOCT0.OCTHRFR3, dbo.ORDIOCT0.OCTHRTO3 " & _
                      " FROM dbo.ORDIOCT0 INNER JOIN dbo.VIAGVVC0 ON dbo.ORDIOCT0.OCTIDOCT = dbo.VIAGVVC0.VVCIDOCT " & _
                      " WHERE     (dbo.VIAGVVC0.VVCIDVVT = ?) AND (dbo.VIAGVVC0.VVCIDADZ = ?) "



    
    '-- DP 15/05/08
    '-- Sostituisce SQL_SMS_GET_LIST_AGG
    SQL_SMS_GET_LIST_AGG_2 = "SELECT VVTIDAM1, VVCIDVVT, VVCIDVVC, VVTFLST3, " & _
                           "        VVCFLST5, VVTHRINC, CASE VVCHRPRE WHEN '01/01/1900' THEN VVCTEMVG ELSE VVCHRPRE END as ORA_PREV, VVCTEMVG  " & _
                           "FROM VIAGVVT0 INNER JOIN VIAGVVC0 ON VVTIDVVT = VVCIDVVT " & _
                           "WHERE (VVTDTINV = ? OR VVTDTINV = ?) AND (VVTANNUL = 0) AND (VVCANNUL = 0) AND " & _
                           "        (VVTFLST3 IS NULL OR VVTFLST3 = 'N' OR VVCFLST5 IS NULL OR VVCFLST5 = 'N') " & _
                           "ORDER BY VVTIDAM1"
        
'    SQL_KNOSOS_GET_EVENTS = "SELECT IDVEH, IDVIAJE, NUMVIAJE, CONVERT( nvarchar, FECHA, 21 ), IDDESC, PROD_RES, " & _
'                                " TEVE, CANTREALCAR, CANTREALDESC, KM_PARCIALES, " & _
'                                " CONVERT( nVarchar, DATEADD(Hour, DATEDIFF(Hour, GETUTCDATE(), GETDATE()), FECHA), 21 )" & _
'                                " FROM KNOSEVE0" & _
'                                " ORDER BY FECHA"
    SQL_OESIA_GET_EVENTS = "SELECT CodigoVehiculo, CodigoViaje, CodigoServicio, CONVERT( nvarchar, FechaEvento, 21 ), CodigoEvento, " & _
                            "       Metros, CONVERT(nVarchar, DATEADD(Hour, DATEDIFF(Hour, GETUTCDATE(), GETDATE()), FechaEvento), 21), IdRegistro, " & _
                            "       Latitud, Longitud " & _
                            "FROM EventosMBFlotas " & _
                            "WHERE IdRegistro > ? " & _
                            "ORDER BY FechaEvento "


    SQL_KNOSOS_DELETE_EVENT = "DELETE KNOSEVE0 WHERE IDVEH = ? AND IDVIAJE = ? AND NUMVIAJE = ?" & _
                                " AND FECHA = { Ts '%fecha%' } AND IDDESC = ?"
    
'    SQL_KNOSOS_GET_KEY_CONSEGNA_PRODOTTO_RESOIL = "SELECT VVPIDVVT, VVPIDVVC, VVPIDVVP" & _
'                                        " FROM VIAGVVT0 INNER JOIN" & _
'                                        " VIAGVVC0 ON VIAGVVT0.VVTIDVVT = VIAGVVC0.VVCIDVVT INNER JOIN" & _
'                                        " VIAGVVP0 ON VIAGVVC0.VVCIDVVT = VIAGVVP0.VVPIDVVT" & _
'                                        " AND VIAGVVC0.VVCIDVVC = VIAGVVP0.VVPIDVVC" & _
'                                        " WHERE VIAGVVT0.VVTNRKVI = ? AND VIAGVVP0.VVPIDKDE = ?"
    SQL_KNOSOS_GET_KEY_CONSEGNA_PRODOTTO_RESOIL = "SELECT VVPIDVVT, VVPIDVVC, VVPIDVVP" & _
                                                "  FROM VIAGVVT0 INNER JOIN VIAGVVC0 ON " & _
                                                "       VIAGVVT0.VVTIDVVT = VIAGVVC0.VVCIDVVT " & _
                                                "  INNER JOIN VIAGVVP0 ON " & _
                                                "       VIAGVVC0.VVCIDVVT = VIAGVVP0.VVPIDVVT AND " & _
                                                "       VIAGVVC0.VVCIDVVC = VIAGVVP0.VVPIDVVC " & _
                                                "  INNER JOIN ORDIOCT0 ON " & _
                                                "       VIAGVVC0.VVCIDOCT = ORDIOCT0.OCTIDOCT " & _
                                                "  WHERE VIAGVVT0.VVTIDDEF = ? AND ORDIOCT0.OCTRIFE1 = ? AND " & _
                                                "       VIAGVVT0.VVTANNUL = 0 AND VIAGVVC0.VVCANNUL = 0 AND ORDIOCT0.OCTANNUL = 0"
    
'    SQL_KNOSOS_GET_KEY_VIAGGIO_RESOIL = "SELECT VVTIDVVT FROM VIAGVVT0 WHERE VVTNRKVI = ?"
    SQL_KNOSOS_GET_KEY_VIAGGIO_RESOIL = "SELECT VVTIDVVT FROM VIAGVVT0 WHERE VVTIDDEF = ? AND VVTANNUL = 0"
    
    SQL_KNOSOS_GET_LAST_EVENT_DATE = "SELECT MAX(SEVDTEVE)" & _
                                        " FROM VIAGSEV0"
    
    SQL_KNOSOS_GET_IMPORTO_PEDAGGIO = "SELECT PEAJE FROM KNOSVIA0" & _
                                    " WHERE NUMVIAJE = ? and ESTA = ?"
    
    SQL_KNOSOS_GET_COUNT_SCARICHI_NONCONCLUSI = "SELECT COUNT(*) FROM VIAGVVP0" & _
                                    " INNER JOIN VIAGVVC0 ON VVPIDVVT = VVCIDVVT AND VVPIDVVC = VVCIDVVC" & _
                                    " WHERE VVPIDVVT = ?" & _
                                    " AND VVPIDVVC = ?" & _
                                    " AND (VVPANNUL = 0 AND VVCANNUL = 0)" & _
                                    " AND VVPFLSTA <> '2'"
                                    
    SQL_KNOSOS_GET_COUNT_SCARICHI_NONCONCLUSI_VVT = "SELECT COUNT(*) FROM VIAGVVP0" & _
                                    " INNER JOIN VIAGVVC0 ON VVPIDVVT = VVCIDVVT AND VVPIDVVC = VVCIDVVC" & _
                                    " WHERE VVPIDVVT = ?" & _
                                    " AND (VVPANNUL = 0 AND VVCANNUL = 0)" & _
                                    " AND VVPFLSTA <> '2'"

    SQL_UPDATE_VIAGGIO_MODIFICATO = "UPDATE VIAGVVP0 SET VVPFLST5 = 'S', VVPFLAGG = 'ME' WHERE VVPIDVVT = ? AND VVPANNUL = 0"
    
    SQL_SELECT_KM_EVENTS = "SELECT SEVIDVVC, MIN(SEVNRKM), MAX(SEVNRKM), MIN(SEVDTEVE)" & _
                            " FROM VIAGSEV0" & _
                            " WHERE  SEVIDVVT = ? AND SEVIDVVC <> ''" & _
                            " GROUP  BY SEVIDVVC" & _
                            " ORDER  BY MIN(SEVNRKM), MIN(SEVDTEVE)"
 
    SQL_GET_KM_EVENT = "SELECT SEVNRKM FROM VIAGSEV0" & _
                                " WHERE SEVIDVVT = ? AND SEVIDEVE = ?"
    
    SQL_GET_SELECTION_EVENTS = "SELECT SEVIDSEV, SEVIDVVT, SEVIDVVC, SEVIDVVP," & _
                            " AVV_B.AVVDESCR, SEVDTEVE, IsNull(vvcderas, adedescr), SEVQTLTC, SEVQTLTS, SEVNRKM," & _
                            " AVV_A.AVVDESCR, SEVFLSAT, SEVIDEVE FROM VIAGSEV0" & _
                            " INNER JOIN ANAGAVV AVV_A ON AVV_A.AVVIDTBL = 'ANAGAAM0' AND" & _
                            " AVV_A.AVVIDFLD = 'AAMFLSAT' AND AVV_A.AVVVALUE = VIAGSEV0.SEVFLSAT" & _
                            " INNER JOIN ANAGAVV AVV_B ON AVV_B.AVVIDTBL = 'VIAGSEV0' AND " & _
                            " AVV_B.AVVIDFLD = 'SEVIDEVE' AND AVV_B.AVVVALUE = VIAGSEV0.SEVIDEVE" & _
                            " INNER JOIN VIAGVVT0 ON VVTIDVVT = SEVIDVVT" & _
                            " LEFT JOIN ANAGADE0 ON ADEIDADE = VVTIDAD1 AND (SEVIDEVE = 'CI' OR SEVIDEVE = 'CF')" & _
                            " LEFT JOIN VIAGVVC0 ON VVCIDVVT = SEVIDVVT AND VVCIDVVC = SEVIDVVC" & _
                            " WHERE SEVIDVVT = '%sevidvvt%' ORDER BY SEVDTEVE DESC"

    SQL_DEUSTO_GET_VIAGGI = "SELECT DISTINCT AAMIDDEU, VVTIDAM1 " & _
                            "  FROM VIAGVVT0 INNER JOIN ANAGAAM0 ON VVTIDAM1 = AAMIDAAM " & _
                            " WHERE (VVTDTINV = ? OR VVTDTINC = ?) AND VVTIDASC = '038' AND VVTFLSTV <> 'P' AND VVTANNUL = 0 AND (AAMIDDEU IS NOT NULL) AND (LEN(AAMIDDEU)>0) "
    SQL_SAT_INSERT_POS = "INSERT INTO VIAGPOS0 " & _
                            "       (POSIDAAM, POSDTPOS, POSHRPOS, POSCOORX, POSCOORY, POSFLTRA, POSNRSTA) " & _
                            "VALUES ("
    SQL_SAT_GET_MAX_POS_TRASM = "SELECT MAX(POSIDPOS) " & _
                                   "  FROM VIAGPOS0 " & _
                                   " WHERE POSFLTRA = 'N'"
    SQL_SAT_GET_POS_TRASM_LIST = "SELECT POSIDAAM, POSDTPOS, POSHRPOS, POSCOORX, POSCOORY " & _
                                    "  FROM VIAGPOS0 " & _
                                    " WHERE POSIDPOS <= ? AND POSFLTRA = 'N'"
    SQL_SAT_UPDATE_POS_STATUS = "UPDATE VIAGPOS0 " & _
                                   "   SET POSFLTRA = 'S' " & _
                                   " WHERE POSIDPOS <= ? AND POSFLTRA = 'N'"
'    SQL_SAT_GET_LAST_POS = "SELECT TOP 1 POSCOORX, POSCOORY " & _
'                           "  FROM VIAGPOS0  INNER JOIN ANAGAAM0 ON AAMIDAAM = POSIDAAM " & _
'                           " WHERE (POSIDAAM = ?) AND (AAMIDDEU IS NOT NULL) AND (AAMIDDEU <> '') AND " & _
'                           " (CAST(DATEPART(yyyy, POSDTPOS) AS CHAR(4)) + RIGHT(CAST(DATEPART(mm, POSDTPOS) + 100 AS CHAR(3)), 2) " & _
'                           "  + RIGHT(CAST(DATEPART(dd, POSDTPOS) + 100 AS CHAR(3)), 2) + RIGHT(CAST(DATEPART(hh, POSHRPOS) + 100 AS CHAR(3)), 2) " & _
'                           "  + RIGHT(CAST(DATEPART(n, POSHRPOS) + 100 AS CHAR(3)), 2) + RIGHT(CAST(DATEPART(ss, POSHRPOS) + 100 AS CHAR(3)), 2) <= ?) " & _
'                           " ORDER BY CAST(DATEPART(yyyy, POSDTPOS) AS CHAR(4)) + RIGHT(CAST(DATEPART(mm, POSDTPOS) + 100 AS CHAR(3)), 2) " & _
'                           "  + RIGHT(CAST(DATEPART(dd, POSDTPOS) + 100 AS CHAR(3)), 2) + RIGHT(CAST(DATEPART(hh, POSHRPOS) + 100 AS CHAR(3)), 2) " & _
'                           "  + RIGHT(CAST(DATEPART(n, POSHRPOS) + 100 AS CHAR(3)), 2) + RIGHT(CAST(DATEPART(ss, POSHRPOS) + 100 AS CHAR(3)), 2) DESC "
    SQL_SAT_GET_LAST_POS = "SELECT POSCOORX, POSCOORY, POSDTPOS, POSHRPOS " & _
                           "  FROM VIAGPOS0  INNER JOIN ANAGAAM0 ON AAMIDAAM = POSIDAAM " & _
                           " WHERE (POSIDAAM = "
                           
    SQL_SAT_GET_LAST_POS_2 = ") AND (POSDTPOS >= "
    
    SQL_SAT_GET_LAST_POS_3 = ") AND (POSDTPOS <= "
                           
    SQL_SAT_GET_LAST_POS_4 = ") ORDER BY POSDTPOS DESC, POSHRPOS DESC "
    
    '-- DP 20/06/08
    '-- Query per controllo SMS doppi
    SQL_FIND_DUBLE_SMS = "SELECT SMSIDSMS From VIAGSMS0 WHERE SMSIDSMS < ? AND SMSNRTEL = ? AND (SMSDTSMD <= ? AND SMSDTSMD >= ?) AND SMSTESTO = ?"
    
    Set DS = New CDataStruct

    With DS.Fields
        .Add "Numero Evento", "VIAGSEV0.sevidsev", 0, True, 7, "0", , True
        .Add "Numero Viaggio", "VIAGSEV0.sevidvvt", 0, True, 7, "1", , True
        .Add "Numero Consegna", "VIAGSEV0.sevidvvc", 0, True, 3, "2", , True
        .Add "Numero Scarico", "VIAGSEV0.sevidvvp", 0, True, 3, "3", , True
        .Add "ID evento", "VIAGSEV0.sevideve", 0, True, 3, "4", , True
        .Add "Data evento", "VIAGSEV0.sevdteve", 2, True, 23, "5", , True
        .Add "Litri caricati", "VIAGSEV0.sevqtltc", 1, False, 3, "6", , True
        .Add "Litri scaricati", "VIAGSEV0.sevqtlts", 1, False, 3, "7", , True
        .Add "Numero KM", "VIAGSEV0.sevnrkm", 1, False, 3, "8", , True
        .Add "Tipo satellitare", "VIAGSEV0.sevflsat", 0, True, 1, "9", , True
    End With
    
    Set ADZ = New CDestinazione
    Dim DataConnect As CConnect
    Set DataConnect = New CConnect
    With ADZ.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    
    With DS.KeyFields
        .Add "0", "1"
        .Add "1", "2"
  End With
    DS.CancelField = ""
    
    Set objLogger = New CLogger

End Sub

Private Sub Class_Terminate()
    Set DS = Nothing
    Set objLogger = Nothing
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
Function StringKNOSOSDate_To_Date(strDate As String)
    If Len(strDate) = 23 Then
        strDate = Left(strDate, Len(strDate) - 4)
    End If
    
    StringKNOSOSDate_To_Date = CDate(strDate)
End Function
Function convertLocalDateToUTCDate(strDate As String) As String
    Dim varBuffer As Variant
    Dim lngHowMany As Long
    Dim strSQLStatement As String

    strSQLStatement = Replace("SELECT CONVERT( nVarchar, DATEADD(hour, -DATEDIFF(hour, GETUTCDATE(), GETDATE()), CONVERT(datetime, { Ts '%fecha%' }, 21)), 21)", "%fecha%", strDate)
    If Not DS.DA.QueryCommand(strSQLStatement, varBuffer, lngHowMany) Then
        convertLocalDateToUTCDate = ""
        Exit Function
    End If
        
    convertLocalDateToUTCDate = varBuffer(0, 0)
    
End Function
Function TruncDate(strDate As Date)
    TruncDate = CDate(DatePart("h", strDate) & ":" & DatePart("n", strDate) & "." & DatePart("s", strDate))
End Function
Private Sub LogFieldValue(Field As CField)

    If Field.Updated Then
        objLogger.logDebug "+ """ & Field.Name & """ (" & Field.FieldName & ") = [" & Field.NewValue & "]"
    Else
        objLogger.logDebug "  """ & Field.Name & """ (" & Field.FieldName & ") = [" & Field.Value & "]"
    End If
End Sub

' -- DP 31/05/08
' -- Commento questa funzione che non viene utilizzata in TPH Siviglia
'Function ScanEvents_KNOSOS() As Boolean
'
'    Dim i As Long
'    Dim varKBuffer As Variant
'    Dim varBuffer As Variant
'    Dim varKBuffer2 As Variant
'    Dim strLastDate As String
'    Dim lngHowMany As Long
'    Dim varIDViaggio As Variant
'    Dim varIDConsegna As Variant
'    Dim varIDProdotto As Variant
'    Dim VVP As CViaggioProdotto
'    Dim VVT As CViaggio
'    Dim VVC As CViaggioConsegna
'    Dim dateLocalEventDateTime As Date
'    Dim sngKm As Single
'    Dim strSQLStatement As String
'    Dim blnUsaDatiGPS As Boolean
'
'    objLogger.logDebug "Inizio ScanEvents_KNOSOS."
'
'    ' Scan KNOSOS table events
'    lngHowMany = 0
'    If Not DS.DA.QueryCommand(SQL_KNOSOS_GET_EVENTS, varKBuffer, lngHowMany) Then
'        objLogger.logError "Query SQL_KNOSOS_GET_EVENTS."
'        GoTo ErrHandler
'    End If
'
'    Set VVT = New CViaggio
'    With VVT
'        Set .DS.DA = DS.DA
'        Set .DS.ErrH = DS.DA.ErrH
'    End With
'
'    Set VVP = New CViaggioProdotto
'    With VVP
'        Set .DS.DA = DS.DA
'        Set .DS.ErrH = DS.DA.ErrH
'    End With
'
'    Set VVC = New CViaggioConsegna
'    With VVC
'        Set .DS.DA = DS.DA
'        Set .DS.ErrH = DS.DA.ErrH
'    End With
'
'    ' Inizia transazione
'    DS.DA.BeginTransaction
'
'    objLogger.logInfo "Trovati " & lngHowMany & " eventi."
'
'    If lngHowMany > 0 Then
'
'        For i = LBound(varKBuffer, 2) To UBound(varKBuffer, 2)
'
'            varIDViaggio = ""
'            varIDConsegna = ""
'            varIDProdotto = ""
'
'            objLogger.logInfo "Evento n. " & i + 1
'            objLogger.logInfo "-------------------------------------------------"
'            objLogger.logInfo "IDVEH         [" & varKBuffer(0, i) & "]"
'            objLogger.logInfo "IDVIAJE       [" & varKBuffer(1, i) & "]"
'            objLogger.logInfo "NUMVIAJE      [" & varKBuffer(2, i) & "]"
'            objLogger.logInfo "FECHA         [" & varKBuffer(3, i) & "]"
'            objLogger.logInfo "IDDESC        [" & varKBuffer(4, i) & "]"
'            objLogger.logInfo "TEVE          [" & varKBuffer(5, i) & "]"
'            objLogger.logInfo "PROD_RES      [" & varKBuffer(6, i) & "]"
'            objLogger.logInfo "CANTREALCAR   [" & varKBuffer(7, i) & "]"
'            objLogger.logInfo "CANTREALDESC  [" & varKBuffer(8, i) & "]"
'            objLogger.logInfo "KM_PARCIALES  [" & varKBuffer(9, i) & "]"
'            objLogger.logInfo "-------------------------------------------------"
'
'            ' E' un evento relativo al viaggio o carico o scarico?
'            If CSng(varKBuffer(4, i)) > 0 Then
'
'                ' E' uno scarico.
'                lngHowMany = 0
'                If Not DS.DA.QueryCommand(SQL_KNOSOS_GET_KEY_CONSEGNA_PRODOTTO_RESOIL, varBuffer, lngHowMany, _
'                    CStr(varKBuffer(2, i)), CSng(varKBuffer(4, i))) Then
'                    objLogger.logError "Query SQL_KNOSOS_GET_KEY_CONSEGNA_PRODOTTO_RESOIL."
'                    GoTo ErrHandler
'                End If
'                If lngHowMany > 0 Then
'                    varIDViaggio = varBuffer(0, 0)
'                    varIDConsegna = varBuffer(1, 0)
'                    varIDProdotto = varBuffer(2, 0)
'                Else
'                    objLogger.logWarning "Consegna prodotto non trovato in ResOil."
'                End If
'            Else
'
'                ' E' un carico o viaggio.
'                lngHowMany = 0
'                If Not DS.DA.QueryCommand(SQL_KNOSOS_GET_KEY_VIAGGIO_RESOIL, varBuffer, lngHowMany, _
'                    CStr(varKBuffer(2, i))) Then
'                    objLogger.logError "Query SQL_KNOSOS_GET_KEY_VIAGGIO_RESOIL."
'                    GoTo ErrHandler
'                End If
'                If lngHowMany > 0 Then
'                    varIDViaggio = varBuffer(0, 0)
'                    varIDConsegna = ""
'                    varIDProdotto = ""
'                Else
'                    objLogger.logWarning "Viaggio non trovato in ResOil."
'                End If
'            End If
'
'            objLogger.logInfo "ResOil Keys VVTIDVVT = """ & varIDViaggio & _
'                                """, VVCIDVVC = """ & varIDConsegna & """, VVPIDVVP = """ & varIDProdotto & """"
'
'            ' Log record event
'            DS.DA.LogExecTime "INFO: KNOSOS Evento", -1
'            DS.DA.LogExecTime "INFO: -----------------------------------------------------------------", -1
'            DS.DA.LogExecTime "INFO: - IDVEH         [" & varKBuffer(0, i) & "]", -1
'            DS.DA.LogExecTime "INFO: - IDVIAJE       [" & varKBuffer(1, i) & "]", -1
'            DS.DA.LogExecTime "INFO: - NUMVIAJE      [" & varKBuffer(2, i) & "]", -1
'            DS.DA.LogExecTime "INFO: - FECHA         [" & varKBuffer(3, i) & "]", -1
'            DS.DA.LogExecTime "INFO: - IDDESC        [" & varKBuffer(4, i) & "]", -1
'            DS.DA.LogExecTime "INFO: - TEVE          [" & varKBuffer(5, i) & "]", -1
'            DS.DA.LogExecTime "INFO: - PROD_RES      [" & varKBuffer(6, i) & "]", -1
'            DS.DA.LogExecTime "INFO: - CANTREALCAR   [" & varKBuffer(7, i) & "]", -1
'            DS.DA.LogExecTime "INFO: - CANTREALDESC  [" & varKBuffer(8, i) & "]", -1
'            DS.DA.LogExecTime "INFO: - KM_PARCIALES  [" & varKBuffer(9, i) & "]", -1
'
'            ' Se non trovo il viaggio o prodotto, skip
'            If varIDViaggio <> "" Then
'
'                dateLocalEventDateTime = StringKNOSOSDate_To_Date(CStr(varKBuffer(10, i)))
'                If lngHowMany > 0 Then
'
'                    DS.Fields.ResetFields
'
'                    ' Creazione dell'evento
'                    DS.Fields("0").NewValue = DS.DA.GetCounter("anaganu0", "anuprogr", "anuidanu", "901", 0)
'                    DS.Fields("0").Updated = True
'
'                    DS.Fields("1").NewValue = varIDViaggio
'                    DS.Fields("1").Updated = True
'
'                    DS.Fields("2").NewValue = varIDConsegna
'                    DS.Fields("2").Updated = True
'
'                    DS.Fields("3").NewValue = varIDProdotto
'                    DS.Fields("3").Updated = True
'
'                    DS.Fields("5").NewValue = varKBuffer(10, i)
'                    DS.Fields("5").Updated = True
'
'                    DS.Fields("6").NewValue = varKBuffer(7, i)
'                    DS.Fields("6").Updated = True
'
'                    DS.Fields("7").NewValue = varKBuffer(8, i)
'                    DS.Fields("7").Updated = True
'
'                    DS.Fields("8").NewValue = varKBuffer(9, i)
'                    DS.Fields("8").Updated = True
'
'                    DS.Fields("9").NewValue = "K"
'                    DS.Fields("9").Updated = True
'
'                    ' Carica il viaggio
'                    With VVT
'                        .DS.Fields.ResetFields
'                        .DS.Fields("0").FindValue = varIDViaggio
'
'                        If Not .Find Then
'                            objLogger.logError "Find Viaggio."
'                            GoTo ErrHandler
'                        End If
'
'                        If IsEmptyVariantString(.DS.Field("Flag Usa Dati GPS").Value) Or _
'                            .DS.Field("Flag Usa Dati GPS").Value = "S" Then
'                            blnUsaDatiGPS = True
'                        Else
'                            blnUsaDatiGPS = False
'                        End If
'
'                    End With
'
'
'                    Select Case varKBuffer(6, i)
'
'                        ' -------------------------------------------------------
'                        ' Evento - Inicio del viaggio (Inizio del viaggio)
'                        ' -------------------------------------------------------
'                        Case "ON "
'                            DS.Fields("4").NewValue = "VI"
'                            DS.Fields("4").Updated = True
'
'                            objLogger.logInfo "Tipo evento: Inizio viaggio"
'
'                            ' Aggiorna lo stato del viaggio (VVTFLSTA)
'                            VVT.DS.Fields("48").NewValue = VIAGGIO_STATO_INIZIOVIAGGIO
'                            VVT.DS.Fields("48").Updated = True
'                            LogFieldValue VVT.DS.Fields("48")
'
'                            If blnUsaDatiGPS Then
'                                ' Aggiorna il flag modificato (VVPFLST5) di tutti i prodotti del viaggio
'                                If Not DS.DA.ExecuteCommand(SQL_UPDATE_VIAGGIO_MODIFICATO, varIDViaggio) Then
'                                    objLogger.logError "Errore nell'aggiornamento del flag Modificato."
'                                    GoTo ErrHandler
'                                End If
'                            End If
'
'                        ' -------------------------------------------------------
'                        ' Evento - Llegada en factora (Inizio carico)
'                        ' -------------------------------------------------------
'                        Case "LLE"
'                            DS.Fields("4").NewValue = "CI"
'                            DS.Fields("4").Updated = True
'
'                            objLogger.logInfo "Tipo evento: Inizio carico"
'
'                            ' Aggiorna lo stato del viaggio (VVTFLSTA)
'                            VVT.DS.Fields("48").NewValue = VIAGGIO_STATO_CARICO
'                            VVT.DS.Fields("48").Updated = True
'                            LogFieldValue VVT.DS.Fields("48")
'
'                            If blnUsaDatiGPS Then
'                                ' Aggiorna l'ora inizio carico sul viaggio (VVTHRINC)
'                                VVT.DS.Fields("5").NewValue = TruncDate(dateLocalEventDateTime)
'                                VVT.DS.Fields("5").Updated = True
'                                LogFieldValue VVT.DS.Fields("5")
'
'                                ' Aggiorna il flag modificato (VVPFLST5) di tutti i prodotti del viaggio
'                                If Not DS.DA.ExecuteCommand(SQL_UPDATE_VIAGGIO_MODIFICATO, varIDViaggio) Then
'                                    objLogger.logError "Errore nell'aggiornamento del flag Modificato."
'                                    GoTo ErrHandler
'                                End If
'                            End If
'
'                        ' -------------------------------------------------------
'                        ' Evento - Salida de factora (Fine del carico)
'                        ' -------------------------------------------------------
'                        Case "SAL"
'                            DS.Fields("4").NewValue = "CF"
'                            DS.Fields("4").Updated = True
'
'                            objLogger.logInfo "Tipo evento: Fine carico"
'
'                            ' Aggiorna lo stato del viaggio (VVTFLSTA)
'                            VVT.DS.Fields("48").NewValue = VIAGGIO_STATO_INCORSO
'                            VVT.DS.Fields("48").Updated = True
'                            LogFieldValue VVT.DS.Fields("48")
'
'                            ' Aggiorna il tempo di carico sul viaggio (VVTTMPCA)
'                            If blnUsaDatiGPS Then
'                                VVT.DS.Fields("30").NewValue = DateDiff("n", TruncDate(VVT.DS.Fields("5").Value), TruncDate(dateLocalEventDateTime))
'                                VVT.DS.Fields("30").Updated = True
'                                LogFieldValue VVT.DS.Fields("30")
'
'                                ' Aggiorna l'ora inizio viaggio sul viaggio (VVTHRINV)
'                                VVT.DS.Fields("7").NewValue = TruncDate(dateLocalEventDateTime)
'                                VVT.DS.Fields("7").Updated = True
'                                LogFieldValue VVT.DS.Fields("7")
'
'                                ' Aggiorna flag carico completato
'                                VVT.DS.Fields("39").NewValue = "S"
'                                VVT.DS.Fields("39").Updated = True
'                                LogFieldValue VVT.DS.Fields("39")
'
'                                ' Aggiorna il flag modificato (VVPFLST5) di tutti i prodotti della consegna
'                                If Not DS.DA.ExecuteCommand(SQL_UPDATE_VIAGGIO_MODIFICATO, varIDViaggio) Then
'                                    objLogger.logError "Errore in aggiornamento flag Modificato del viaggio."
'                                    GoTo ErrHandler
'                                End If
'                            End If
'
'                        ' -------------------------------------------------------
'                        ' Evento - Inicio descarga (Inizio scarico prodotto)
'                        ' -------------------------------------------------------
'                        Case "INI"
'                            DS.Fields("4").NewValue = "SI"
'                            DS.Fields("4").Updated = True
'
'                            objLogger.logInfo "Tipo evento: Inizio consegna prodotto"
'
'                            ' Carica la consegna prodotto e aggiorna i dati
'                            With VVP
'                                .DS.Fields.ResetFields
'                                .DS.Fields("0").FindValue = varIDViaggio
'                                .DS.Fields("1").FindValue = varIDConsegna
'                                .DS.Fields("2").FindValue = varIDProdotto
'
'                                If Not .Find Then
'                                    objLogger.logError "Find Consegna Prodotto."
'                                    GoTo ErrHandler
'                                End If
'
'                                ' Aggiorna lo stato dello scarico (VVPFLSTA)
'                                .DS.Fields("37").NewValue = "1"
'                                .DS.Fields("37").Updated = True
'                                LogFieldValue .DS.Fields("37")
'
'                                ' Aggiorna il flag Modificato (VVPFLST5)
'                                .DS.Field("Stato Riga 5").NewValue = "S"
'                                .DS.Field("Stato Riga 5").Updated = True
'                                LogFieldValue .DS.Field("Stato Riga 5")
'
'                                ' Aggiorna l'ora inizio scarico (VVPHRINC)
'                                If blnUsaDatiGPS Then
'                                    .DS.Fields("38").NewValue = TruncDate(dateLocalEventDateTime)
'                                    .DS.Fields("38").Updated = True
'                                    LogFieldValue .DS.Fields("38")
'                                End If
'
'                                ' Salva
'                                If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
'                                    objLogger.logError "Save Consegna Prodotto."
'                                    GoTo ErrHandler
'                                End If
'
'                                .DS.DA.EndTrans etCommit
'                            End With
'
'                            ' Carica la consegna e aggiorna i dati
'                            With VVC
'
'                                .DS.Fields.ResetFields
'                                .DS.Fields("0").FindValue = varIDViaggio
'                                .DS.Fields("1").FindValue = varIDConsegna
'
'                                If Not .Find Then
'                                    objLogger.logError "Find Consegna."
'                                    GoTo ErrHandler
'                                End If
'
'                                ' E' il primo scarico della consegna
'                                If .DS.Fields("44").Value = "0" Or IsNull(.DS.Fields("44").Value) Or .DS.Fields("44").Value = "" Then
'
'                                    If blnUsaDatiGPS Then
'                                        '-- DP 15/05/08
'                                        '-- Aggiorna l'ora di arrivo sul nuovo campo VVCHRARR anziché VVCTEMVG
'                                        .DS.Fields("47").NewValue = TruncDate(dateLocalEventDateTime)
'                                        .DS.Fields("47").Updated = True
'                                        LogFieldValue .DS.Fields("47")
'                                    End If
'
'                                    ' Aggiorna lo stato 'In corso'
'                                    .DS.Fields("44").NewValue = "1"
'                                    .DS.Fields("44").Updated = True
'                                    LogFieldValue .DS.Fields("44")
'                                End If
'
''                                ' Imposta i km parziali
''                                If Not (IsNull(varKBuffer(9, i)) Or IsEmpty(varKBuffer(9, i))) Then
''                                    .DS.Fields("22").NewValue = GetKmConsegna(varIDViaggio, varIDConsegna) + CSng(varKBuffer(9, i))
''                                    .DS.Fields("22").Updated = True
''                                End If
'
'                                ' Salva la consegna
'                                If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
'                                    objLogger.logError "Save Consegna."
'                                    GoTo ErrHandler
'                                End If
'
'                                .DS.DA.EndTrans etCommit
'                            End With
'
'                        ' -------------------------------------------------------
'                        ' Evento - Fin descarga (Fine scarico prodotto)
'                        ' -------------------------------------------------------
'                        Case "FIN"
'                            DS.Fields("4").NewValue = "SF"
'                            DS.Fields("4").Updated = True
'
'                            objLogger.logInfo "Tipo evento: Fine consegna prodotto"
'
'                            ' Carica la consegna prodotto e aggiorna i dati
'                            With VVP
'                                .DS.Fields.ResetFields
'                                .DS.Fields("0").FindValue = varIDViaggio
'                                .DS.Fields("1").FindValue = varIDConsegna
'                                .DS.Fields("2").FindValue = varIDProdotto
'
'                                If Not .Find Then
'                                    objLogger.logError "Find Consegna Prodotto."
'                                    GoTo ErrHandler
'                                End If
'
'                                ' Aggiorna lo stato dello scarico (VVPFLSTA)
'                                .DS.Fields("37").NewValue = "2"
'                                .DS.Fields("37").Updated = True
'                                LogFieldValue .DS.Fields("37")
'
'                                If blnUsaDatiGPS Then
'                                    ' Imposta i litri scaricati
'                                    If IsNull(varKBuffer(8, i)) Or IsEmpty(varKBuffer(8, i)) Then
'                                        .DS.Fields("11").NewValue = 0
'                                    Else
'                                        .DS.Fields("11").NewValue = CSng(varKBuffer(8, i))
'                                    End If
'                                    .DS.Fields("11").Updated = True
'                                    LogFieldValue .DS.Fields("11")
'
'                                    ' Imposta i litri caricati
'                                    .DS.Fields("7").NewValue = .DS.Fields("11").NewValue
'                                    .DS.Fields("7").Updated = True
'                                    LogFieldValue .DS.Fields("7")
'
'                                    ' Aggiorna l'ora fine scarico (VVPHRFIC)
'                                    .DS.Fields("39").NewValue = TruncDate(dateLocalEventDateTime)
'                                    .DS.Fields("39").Updated = True
'                                    LogFieldValue .DS.Fields("39")
'
'                                    ' Aggiorna il flag Modificato (VVPFLST5)
'                                    .DS.Field("Stato Riga 5").NewValue = "S"
'                                    .DS.Field("Stato Riga 5").Updated = True
'                                    LogFieldValue .DS.Field("Stato Riga 5")
'                                End If
'
'                                ' Salva
'                                If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
'                                    objLogger.logError "Save Consegna Prodotto."
'                                    GoTo ErrHandler
'                                End If
'
'                                .DS.DA.EndTrans etCommit
'                            End With
'
'                            ' Carica la consegna e aggiorna i dati
'                            With VVC
'
'                                .DS.Fields.ResetFields
'                                .DS.Fields("0").FindValue = varIDViaggio
'                                .DS.Fields("1").FindValue = varIDConsegna
'
'                                If Not .Find Then
'                                    objLogger.logError "Find Consegna."
'                                    GoTo ErrHandler
'                                End If
'
'                                ' Se e' l'ultimo scarico della consegna, aggiorna il tempo scarico e lo stato
'                                lngHowMany = 0
'                                If Not DS.DA.QueryCommand(SQL_KNOSOS_GET_COUNT_SCARICHI_NONCONCLUSI, varBuffer, lngHowMany, _
'                                    varIDViaggio, varIDConsegna) Then
'                                    objLogger.logError "Query SQL_KNOSOS_GET_COUNT_SCARICHI_NONCONCLUSI."
'                                    GoTo ErrHandler
'                                End If
'
'                                If CLng(varBuffer(0, 0)) = 0 Then
'
'                                    If blnUsaDatiGPS Then
'                                        '-- DP 15/05/08
'                                        '-- Aggiorna il tempo di scarico utilizzando il nuovo campo VVCHRARR anziché VVCTEMVG
'                                        .DS.Fields("25").NewValue = DateDiff("n", TruncDate(.DS.Fields("47").Value), TruncDate(dateLocalEventDateTime))
'                                        .DS.Fields("25").Updated = True
'                                        LogFieldValue .DS.Fields("25")
'
'                                        ' Aggiorna il flag consegna ultimata
'                                        .DS.Fields("36").NewValue = "S"
'                                        .DS.Fields("36").Updated = True
'                                        LogFieldValue .DS.Fields("36")
'                                    End If
'
'                                    ' Aggiorna lo stato 'Completo'
'                                    .DS.Fields("44").NewValue = "2"
'                                    .DS.Fields("44").Updated = True
'                                    LogFieldValue .DS.Fields("44")
'                                End If
'
''                                ' Imposta i km parziali
''                                If Not (IsNull(varKBuffer(9, i)) Or IsEmpty(varKBuffer(9, i))) Then
''                                    .DS.Fields("22").NewValue = GetKmConsegna(varIDViaggio, varIDConsegna) + CSng(varKBuffer(9, i))
''                                    .DS.Fields("22").Updated = True
''                                End If
'
'                                ' Salva la consegna
'                                If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
'                                    objLogger.logError "Save Consegna."
'                                    GoTo ErrHandler
'                                End If
'
'                                .DS.DA.EndTrans etCommit
'                            End With
'
'                        ' -------------------------------------------------------
'                        ' Evento - Rechazo de descarga (Dirottamento scarico)
'                        ' -------------------------------------------------------
'                        Case "REC"
'                            DS.Fields("4").NewValue = "SA"
'                            DS.Fields("4").Updated = True
'
'                            objLogger.logInfo "Tipo evento: Dirottamento o Rientro"
'
'                            ' Carica la consegna
'                            With VVC
'
'                                .DS.Fields.ResetFields
'                                .DS.Fields("0").FindValue = varIDViaggio
'                                .DS.Fields("1").FindValue = varIDConsegna
'
'                                If Not .Find Then
'                                    objLogger.logError "Find Consegna."
'                                    GoTo ErrHandler
'                                End If
'
'                            End With
'
'                            ' Carica la consegna prodotto e aggiorna i dati
'                            With VVP
'                                .DS.Fields.ResetFields
'                                .DS.Fields("0").FindValue = varIDViaggio
'                                .DS.Fields("1").FindValue = varIDConsegna
'                                .DS.Fields("2").FindValue = varIDProdotto
'
'                                If Not .Find Then
'                                    objLogger.logError "Find Consegna Prodotto."
'                                    GoTo ErrHandler
'                                End If
'
'                                ' Aggiorna lo stato dello scarico (VVPFLSTA)
'                                .DS.Fields("37").NewValue = "2"
'                                .DS.Fields("37").Updated = True
'                                LogFieldValue .DS.Fields("37")
'
'                                ' Imposta i litri caricati e scaricati
'                                Select Case VVC.DS.Fields("33").Value
'
'                                    ' Dirottamento
'                                    Case "D"
'                                        objLogger.logInfo "Tipo evento: Dirottamento"
'                                        If IsNull(varKBuffer(8, i)) Or IsEmpty(varKBuffer(8, i)) Then
'                                            .DS.Fields("11").NewValue = 0
'                                        Else
'                                            .DS.Fields("11").NewValue = CSng(varKBuffer(8, i))
'                                        End If
'                                        .DS.Fields("11").Updated = True
'                                        LogFieldValue .DS.Fields("11")
'
'                                        ' Imposta i litri caricati
'                                        If blnUsaDatiGPS Then
'                                            .DS.Fields("7").NewValue = 0
'                                            .DS.Fields("7").Updated = True
'                                            LogFieldValue .DS.Fields("7")
'                                        End If
'
'                                    ' Rientro
'                                    Case "R"
'                                        objLogger.logInfo "Tipo evento: Rientro"
'
'                                        If blnUsaDatiGPS Then
'                                            ' Imposta i litri scaricati
'                                            If IsNull(varKBuffer(7, i)) Or IsEmpty(varKBuffer(7, i)) Then
'                                                .DS.Fields("11").NewValue = 0
'                                            Else
'                                                .DS.Fields("11").NewValue = CSng(varKBuffer(7, i))
'                                            End If
'                                            .DS.Fields("11").Updated = True
'                                            LogFieldValue .DS.Fields("11")
'
'                                            ' Imposta i litri caricati
'                                            .DS.Fields("7").NewValue = .DS.Fields("11").NewValue
'                                            .DS.Fields("7").Updated = True
'                                            LogFieldValue .DS.Fields("7")
'                                        End If
'
'                                End Select
'
'                                If blnUsaDatiGPS Then
'                                    ' Aggiorna l'ora fine scarico (VVPHRFIC)
'                                    .DS.Fields("39").NewValue = TruncDate(dateLocalEventDateTime)
'                                    .DS.Fields("39").Updated = True
'                                    LogFieldValue .DS.Fields("39")
'
'                                    ' Aggiorna il flag Modificato (VVPFLST5)
'                                    .DS.Field("Stato Riga 5").NewValue = "S"
'                                    .DS.Field("Stato Riga 5").Updated = True
'                                    LogFieldValue .DS.Field("Stato Riga 5")
'                                End If
'
'                                ' Salva
'                                If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
'                                    objLogger.logError "Save Consegna Prodotto."
'                                    GoTo ErrHandler
'                                End If
'
'                                .DS.DA.EndTrans etCommit
'                            End With
'
'                            ' Carica la consegna e aggiorna i dati
'                            With VVC
'
'                                ' Se e' l'ultimo scarico della consegna, aggiorna il tempo scarico e lo stato
'                                lngHowMany = 0
'                                If Not DS.DA.QueryCommand(SQL_KNOSOS_GET_COUNT_SCARICHI_NONCONCLUSI, varBuffer, lngHowMany, _
'                                    varIDViaggio, varIDConsegna) Then
'                                    objLogger.logError "Query SQL_KNOSOS_GET_COUNT_SCARICHI_NONCONCLUSI."
'                                    GoTo ErrHandler
'                                End If
'
'                                If CLng(varBuffer(0, 0)) = 0 Then
'
'                                    If blnUsaDatiGPS Then
'                                        '-- DP 15/05/08
'                                        '-- Aggiorna il tempo di scarico utilizzando il nuovo campo VVCHRARR anziché VVCTEMVG
'                                        .DS.Fields("25").NewValue = DateDiff("n", TruncDate(.DS.Fields("47").Value), TruncDate(dateLocalEventDateTime))
'                                        .DS.Fields("25").Updated = True
'                                        LogFieldValue .DS.Fields("25")
'
'                                        ' Aggiorna il flag consegna ultimata
'                                        .DS.Fields("36").NewValue = "S"
'                                        .DS.Fields("36").Updated = True
'                                        LogFieldValue .DS.Fields("36")
'                                    End If
'
'                                    ' Aggiorna lo stato 'Completo'
'                                    .DS.Fields("44").NewValue = "2"
'                                    .DS.Fields("44").Updated = True
'                                    LogFieldValue .DS.Fields("44")
'                                End If
'
''                                ' Imposta i km parziali
''                                If Not (IsNull(varKBuffer(9, i)) Or IsEmpty(varKBuffer(9, i))) Then
''                                    .DS.Fields("22").NewValue = GetKmConsegna(varIDViaggio, varIDConsegna) + CSng(varKBuffer(9, i))
''                                    .DS.Fields("22").Updated = True
''                                End If
'
'                                ' Salva la consegna
'                                If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
'                                    objLogger.logError "Save Consegna."
'                                    GoTo ErrHandler
'                                End If
'
'                                .DS.DA.EndTrans etCommit
'                            End With
'
'                        ' -------------------------------------------------------
'                        ' Evento - Devolucin de la descarga (Rientro scarico)
'                        ' -------------------------------------------------------
'                        Case "DEV"
'                            DS.Fields("4").NewValue = "SD"
'                            DS.Fields("4").Updated = True
'
'                            objLogger.logInfo "Tipo evento: Rientro"
'
'                            ' Carica la consegna prodotto e aggiorna i dati
'                            With VVP
'                                .DS.Fields.ResetFields
'                                .DS.Fields("0").FindValue = varIDViaggio
'                                .DS.Fields("1").FindValue = varIDConsegna
'                                .DS.Fields("2").FindValue = varIDProdotto
'
'                                If Not .Find Then
'                                    objLogger.logError "Find Consegna Prodotto."
'                                    GoTo ErrHandler
'                                End If
'
'                                ' Aggiorna lo stato dello scarico (VVPFLSTA)
'                                .DS.Fields("37").NewValue = "2"
'                                .DS.Fields("37").Updated = True
'                                LogFieldValue .DS.Fields("37")
'
'                                If blnUsaDatiGPS Then
'                                    ' Imposta i litri scaricati
'                                    If IsNull(varKBuffer(7, i)) Or IsEmpty(varKBuffer(7, i)) Then
'                                        .DS.Fields("11").NewValue = 0
'                                    Else
'                                        .DS.Fields("11").NewValue = CSng(varKBuffer(7, i))
'                                    End If
'                                    .DS.Fields("11").Updated = True
'                                    LogFieldValue .DS.Fields("11")
'
'                                    ' Imposta i litri caricati
'                                    .DS.Fields("7").NewValue = .DS.Fields("11").NewValue
'                                    .DS.Fields("7").Updated = True
'                                    LogFieldValue .DS.Fields("7")
'
'                                    ' Aggiorna l'ora fine scarico (VVPHRFIC)
'                                    .DS.Fields("39").NewValue = TruncDate(dateLocalEventDateTime)
'                                    .DS.Fields("39").Updated = True
'                                    LogFieldValue .DS.Fields("39")
'
'                                    ' Aggiorna il flag Modificato (VVPFLST5)
'                                    .DS.Field("Stato Riga 5").NewValue = "S"
'                                    .DS.Field("Stato Riga 5").Updated = True
'                                    LogFieldValue .DS.Field("Stato Riga 5")
'                                End If
'
'                                ' Salva
'                                If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
'                                    objLogger.logError "Save Consegna Prodotto."
'                                    GoTo ErrHandler
'                                End If
'
'                                .DS.DA.EndTrans etCommit
'                            End With
'
'                            ' Carica la consegna e aggiorna i dati
'                            With VVC
'
'                                .DS.Fields.ResetFields
'                                .DS.Fields("0").FindValue = varIDViaggio
'                                .DS.Fields("1").FindValue = varIDConsegna
'
'                                If Not .Find Then
'                                    objLogger.logError "Find Consegna."
'                                    GoTo ErrHandler
'                                End If
'
'                                ' Se e' l'ultimo scarico della consegna, aggiorna il tempo scarico e lo stato
'                                lngHowMany = 0
'                                If Not DS.DA.QueryCommand(SQL_KNOSOS_GET_COUNT_SCARICHI_NONCONCLUSI, varBuffer, lngHowMany, _
'                                    varIDViaggio, varIDConsegna) Then
'                                    objLogger.logError "Query SQL_KNOSOS_GET_COUNT_SCARICHI_NONCONCLUSI."
'                                    GoTo ErrHandler
'                                End If
'
'                                If CLng(varBuffer(0, 0)) = 0 Then
'
'                                    If blnUsaDatiGPS Then
'                                        '-- DP 15/05/08
'                                        '-- Aggiorna il tempo di scarico utilizzando il nuovo campo VVCHRARR anziché VVCTEMVG
'                                        .DS.Fields("25").NewValue = DateDiff("n", TruncDate(.DS.Fields("47").Value), TruncDate(dateLocalEventDateTime))
'                                        .DS.Fields("25").Updated = True
'                                        LogFieldValue .DS.Fields("25")
'
'                                        ' Aggiorna il flag consegna ultimata
'                                        .DS.Fields("36").NewValue = "S"
'                                        .DS.Fields("36").Updated = True
'                                        LogFieldValue .DS.Fields("36")
'                                    End If
'
'                                    ' Aggiorna lo stato 'Completo'
'                                    .DS.Fields("44").NewValue = "2"
'                                    .DS.Fields("44").Updated = True
'                                    LogFieldValue .DS.Fields("44")
'                                End If
'
''                                ' Imposta i km parziali
''                                If Not (IsNull(varKBuffer(9, i)) Or IsEmpty(varKBuffer(9, i))) Then
''                                    .DS.Fields("22").NewValue = GetKmConsegna(varIDViaggio, varIDConsegna) + CSng(varKBuffer(9, i))
''                                    .DS.Fields("22").Updated = True
''                                End If
'
'                                ' Salva la consegna
'                                If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
'                                    objLogger.logError "Save Consegna."
'                                    GoTo ErrHandler
'                                End If
'
'                                .DS.DA.EndTrans etCommit
'                            End With
'
'                        ' -------------------------------------------------------
'                        ' Evento - Fin de viaje (Fine del viaggio)
'                        ' -------------------------------------------------------
'                        Case "OFF"
'                            Dim sngImportoPedaggio As Single
'                            Dim lngKmInizioViaggio As Long
'                            Dim lngKmInizioCarico As Long
'
'                            objLogger.logInfo "Tipo evento: Fine viaggio"
'
'                            DS.Fields("4").NewValue = "VF"
'                            DS.Fields("4").Updated = True
'
'                            With VVT
'
'                                ' Aggiorna lo stato del viaggio (VVTFLSTA)
'                                .DS.Fields("48").NewValue = VIAGGIO_STATO_CONCLUSO
'                                .DS.Fields("48").Updated = True
'                                LogFieldValue .DS.Fields("48")
'
'                                If blnUsaDatiGPS Then
'                                    ' Aggiorna l'ora fine viaggio (VVTHRFIV)
'                                    .DS.Fields("9").NewValue = TruncDate(dateLocalEventDateTime)
'                                    .DS.Fields("9").Updated = True
'                                    LogFieldValue .DS.Fields("9")
'
'                                    ' Aggiorna il tempo del viaggio sul viaggio (VVTTMPVI)
'                                    .DS.Fields("29").NewValue = DateDiff("n", TruncDate(.DS.Fields("7").Value), TruncDate(dateLocalEventDateTime))
'                                    .DS.Fields("29").Updated = True
'                                    LogFieldValue .DS.Fields("29")
'
'                                    ' Preleva i km di fine carico
'                                    lngHowMany = 0
'                                    If Not DS.DA.QueryCommand(SQL_GET_KM_EVENT, varBuffer, lngHowMany, varIDViaggio, "CF") Then
'                                        objLogger.logError "Query SQL_GET_KM_EVENT."
'                                        GoTo ErrHandler
'                                    End If
'
'                                    If lngHowMany = 0 Then
'                                        objLogger.logError "Impossibile calcolare il km definitivi del viaggio. Manca l'evento fine carico."
'                                    Else
'                                        ' Aggiorna i km definitivi sul viaggio
'                                        If Not (IsNull(varKBuffer(9, i)) Or IsEmpty(varKBuffer(9, i))) Then
'                                            .DS.Fields("27").NewValue = CLng(varKBuffer(9, i)) - varBuffer(0, 0)
'                                            .DS.Fields("27").Updated = True
'                                        End If
'                                        LogFieldValue .DS.Fields("27")
'                                    End If
'
'                                    ' Preleva i km di inizio viaggio
'                                    lngHowMany = 0
'                                    If Not DS.DA.QueryCommand(SQL_GET_KM_EVENT, varBuffer, lngHowMany, varIDViaggio, "VI") Then
'                                        objLogger.logError "Query SQL_GET_KM_EVENT."
'                                        GoTo ErrHandler
'                                    End If
'
'                                    If lngHowMany = 0 Then
'                                        objLogger.logError "Impossibile aggiornare i km a vuoto. Manca l'evento inizio viaggio."
'                                    Else
'                                        lngKmInizioViaggio = varBuffer(0, 0)
'
'                                        ' Preleva i km di inizio carico
'                                        lngHowMany = 0
'                                        If Not DS.DA.QueryCommand(SQL_GET_KM_EVENT, varBuffer, lngHowMany, varIDViaggio, "CI") Then
'                                            objLogger.logError "Query SQL_GET_KM_EVENT."
'                                            GoTo ErrHandler
'                                        End If
'
'                                        If lngHowMany = 0 Then
'                                            objLogger.logError "Impossibile aggiornare i km a vuoto. Manca l'evento inizio carico."
'                                        Else
'                                            lngKmInizioCarico = varBuffer(0, 0)
'
'                                            ' Aggiorna i km a vuoto
'                                            .DS.Fields("49").NewValue = lngKmInizioCarico - lngKmInizioViaggio
'                                            .DS.Fields("49").Updated = True
'                                            LogFieldValue .DS.Fields("49")
'                                        End If
'                                    End If
'
'                                    ' Aggiorna il flag modificato (VVPFLST5) di tutti i prodotti del viaggio
'                                    If Not DS.DA.ExecuteCommand(SQL_UPDATE_VIAGGIO_MODIFICATO, varIDViaggio) Then
'                                        objLogger.logError "Errore in aggiornamento flag Modificato del viaggio."
'                                        GoTo ErrHandler
'                                    End If
'                                End If
'
''                                ' Aggiorna i km totali del viaggio
''                                .DS.Fields("49").NewValue = CLng(varKBuffer(9, i)) - lngKmInizioViaggio
''                                .DS.Fields("49").Updated = True
''                                LogFieldValue .DS.Fields("49")
'
'                            End With
'
'                            If blnUsaDatiGPS Then
'
'                                ' Preleva l'importo pedaggio
'                                lngHowMany = 0
'                                If Not DS.DA.QueryCommand(SQL_KNOSOS_GET_IMPORTO_PEDAGGIO, varKBuffer2, lngHowMany, _
'                                                CStr(varKBuffer(2, i)), KNOSOS_VIAGGIO_STATO_TERMINATO) Then
'                                    objLogger.logError "Query SQL_KNOSOS_GET_IMPORTO_PEDAGGIO."
'                                    GoTo ErrHandler
'                                End If
'
'                                If lngHowMany > 0 Then
'                                    sngImportoPedaggio = CSng(Replace(varKBuffer2(0, 0), ".", ","))
'                                Else
'                                    sngImportoPedaggio = 0
'                                End If
'
'                                ' Calcola imponibile pedaggio
'                                If sngImportoPedaggio > 0 Then
'                                    sngImportoPedaggio = Round(sngImportoPedaggio / (1 + (ALIQUOTA_IVA_PEDAGGIO / 100)), 2)
'                                End If
'
'                                ' Aggiorna il campo importo pedaggio
'                                VVT.DS.Fields("46").NewValue = sngImportoPedaggio
'                                VVT.DS.Fields("46").Updated = True
'                                LogFieldValue VVT.DS.Fields("46")
'
'                                ' Aggiorna i km
'                                If Not UpdateKM(varIDViaggio, CSng(varKBuffer(9, i))) Then
'                                    GoTo ErrHandler
'                                End If
'                            End If
'
'                        Case Else
'                            objLogger.logError "Tipo evento sconosciuto."
'                            GoTo ErrHandler
'
'                    End Select
'
'                    ' Aggiorna il viaggio
'                    With VVT
''                        If Not (IsNull(varKBuffer(9, i)) Or IsEmpty(varKBuffer(9, i))) Then
''                            .DS.Fields("27").NewValue = GetKmViaggio(varIDViaggio) + CSng(varKBuffer(9, i))
''                            .DS.Fields("27").Updated = True
''                        End If
'
'                        If blnUsaDatiGPS Then
'                            ' Aggiorna il flag trasmissione viaggio
'                            .DS.Fields("41").NewValue = "S"
'                            .DS.Fields("41").Updated = True
'                            LogFieldValue .DS.Fields("41")
'                        End If
'
'                        ' Salva
'                        If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
'                            objLogger.logError "Save Viaggio."
'                            GoTo ErrHandler
'                        End If
'
'                        .DS.DA.EndTrans etCommit
'                    End With
'
'                    ' Elimina l'evento KNOSOS
'                    objLogger.logInfo "Cancellazione KNOSOS EVENTO."
'                    strSQLStatement = Replace(SQL_KNOSOS_DELETE_EVENT, "%fecha%", varKBuffer(3, i))
'
'                    If Not DS.DA.ExecuteCommand(strSQLStatement, CStr(varKBuffer(0, i)), _
'                                                CStr(varKBuffer(1, i)), _
'                                                CStr(varKBuffer(2, i)), _
'                                                CLng(varKBuffer(4, i))) Then
'                        objLogger.logError "Cancellazione Evento KNOSOS."
'                        GoTo ErrHandler
'                    End If
'
'                    objLogger.logInfo "Creazione dell'evento in ResOil."
'                    If Not Insert Then
'                        objLogger.logError "Insert ResOil evento."
'                        GoTo ErrHandler
'                    End If
'
'                End If
'            Else
'                DS.DA.LogExecTime "WARN: Viaggio non trovato!", -1
'            End If
'        Next i
'    End If
'
'    ScanEvents_KNOSOS = True
'
'    DS.DA.EndTrans etCommit
'
'    Set VVP = Nothing
'    Set VVT = Nothing
'    Set VVC = Nothing
'
'    objLogger.logDebug "Fine ScanEvents_KNOSOS senza errori."
'    Exit Function
'
'ErrHandler:
'
'    DS.DA.EndTrans etRollback
'
'    Set VVP = Nothing
'    Set VVT = Nothing
'    Set VVC = Nothing
'
'    objLogger.logError "Fine ScanEvents_KNOSOS con ERRORI."
'
'    ScanEvents_KNOSOS = False
'
'End Function
Function ScanEvents_DEUSTO() As Boolean

    Dim i As Long
    Dim varKBuffer As Variant
    Dim varBuffer As Variant
    Dim varKBuffer2 As Variant
    Dim strLastDate As String
    Dim lngHowMany As Long
    Dim varIDViaggio As Variant
    Dim varIDConsegna As Variant
    Dim varIDProdotto As Variant
    Dim VVP As CViaggioProdotto
    Dim VVT As CViaggio
    Dim VVC As CViaggioConsegna
    Dim dateLocalEventDateTime As Date
    Dim sngKm As Single
    Dim strSQLStatement As String
    Dim blnUsaDatiGPS As Boolean
    
    

    objLogger.logDebug "Inizio ScanEvents_DEUSTO."

    ' Cerchiamo i viaggi che circolano oggi
    lngHowMany = 0
    If Not DS.DA.QueryCommand(SQL_DEUSTO_GET_VIAGGI, varKBuffer, lngHowMany, Date, Date) Then
        objLogger.logError "Query SQL_DEUSTO_GET_VIAGGI."
        GoTo ErrHandler
    End If

    Set VVT = New CViaggio
    With VVT
        Set .DS.DA = DS.DA
        Set .DS.ErrH = DS.DA.ErrH
    End With

    Set VVP = New CViaggioProdotto
    With VVP
        Set .DS.DA = DS.DA
        Set .DS.ErrH = DS.DA.ErrH
    End With

    Set VVC = New CViaggioConsegna
    With VVC
        Set .DS.DA = DS.DA
        Set .DS.ErrH = DS.DA.ErrH
    End With

    objLogger.logInfo "Trovati " & lngHowMany & " automezzi."

    Dim oservice As Object
    Set oservice = CreateObject("VLocalize.VLocalize")
    Dim serviceLoc As String
    Dim timeOut As Integer
    
    ' -- DP 10/02/2009
    'serviceLoc = "http://www.tphweb.es/mbtph/mbws/mbws.asmx"
    serviceLoc = "http://mb.oesia.com/mbflotastph/mbws/mbws.asmx"
    
    timeOut = 5 'Timeout webService in sec.

    'Parametri WebService
    Dim user As String
    Dim Password As String
    Dim idVehicle As Integer

    user = "wstph"
    Password = "swtph"

    If lngHowMany > 0 Then

        For i = LBound(varKBuffer, 2) To UBound(varKBuffer, 2)

            If IsNumeric(varKBuffer(0, i)) Then
            idVehicle = varKBuffer(0, i)

            'Parametri Restituiti
            Dim x As String
            Dim Y As String
            Dim ErrCode As String
            Dim ErrDescription As String
            Dim PosDate As String
            Dim PosHour As String
            Dim Status As String

                objLogger.logInfo "Veicolo " & idVehicle
                If (oservice.GetVehiclePos(user, Password, idVehicle, serviceLoc, timeOut, x, Y, PosDate, PosHour, Status, ErrCode, ErrDescription)) Then
                    objLogger.logInfo "Coordinate X:" + x + " |Y:" + Y + " |PosDate:" + PosDate + " |PosHour:" + PosHour + " |Stato:" + Status
                    strSQLStatement = SQL_SAT_INSERT_POS & DS.DA.MakeSQL(varKBuffer(1, i), dtAlfa) & ","
                    strSQLStatement = strSQLStatement & DS.DA.MakeSQL(PosDate, dtDate) & ","
                    strSQLStatement = strSQLStatement & "'" & Replace(Format(PosHour, "hh:mm:ss"), ".", ":") & "',"
                    strSQLStatement = strSQLStatement & DS.DA.MakeSQL(x, dtNum) & ","
                    strSQLStatement = strSQLStatement & DS.DA.MakeSQL(Y, dtNum) & ","
                    strSQLStatement = strSQLStatement & DS.DA.MakeSQL("N", dtAlfa) & ","
                    strSQLStatement = strSQLStatement & DS.DA.MakeSQL(Status, dtNum) & ")"
                    
                    If DS.DA.ExecuteQuery(strSQLStatement) Then
                        objLogger.logInfo "Posizione veicolo inserita correttamente"
                    Else
                        objLogger.logInfo "Errore durante l'inserimento della posizione del veicolo"
                    End If
                Else
                    objLogger.logError "Errore: " & ErrCode & " - " & Err.Description
                End If
            Else
            End If

        Next i
    End If

    ScanEvents_DEUSTO = True

    objLogger.logDebug "Fine ScanEvents_DEUSTO senza errori."
    Exit Function

ErrHandler:

    DS.DA.EndTrans etRollback

    objLogger.logError "Fine ScanEvents_DEUSTO con ERRORI."

    ScanEvents_DEUSTO = False

End Function

Function ScanEvents_Palm() As Boolean
    Dim i As Long
    Dim j As Long
    Dim varKBuffer As Variant
    Dim varBuffer As Variant
    Dim varKBuffer2 As Variant
    Dim varkbufferord As Variant
    Dim varkbuffercliord As Variant
    Dim varkbufferordcli As Variant
    Dim varBufferConsegna As Variant
    Dim varDataEvento As Variant
    Dim varOraEvento As Variant
    Dim varTelAutista As Variant
    Dim strLastDate As String
    Dim lngHowMany As Long
    Dim lngHowManyordcli As Long
    Dim lngHowManyordcli2 As Long
    Dim lnghowmanycons As Long
    Dim varIDViaggio As Variant
    Dim varIDConsegna As Variant
    Dim varIDProdotto As Variant
    Dim varIDAutista As Variant
    Dim varIDSMS As Variant
    Dim varTestoSMS As Variant
    Dim VVP As CViaggioProdotto
    Dim VVT As CViaggio
    Dim VVC As CViaggioConsegna
    Dim AAU As CAutista
    Dim dateLocalEventDateTime As Date
    Dim sngKm As Single
    Dim strSQLStatement As String
    Dim strTipoEvento As String
    Dim blnUsaDatiGPS As Boolean
    Dim varIdevento As String
    Dim IdRegistro As String
    Dim eventooesia As String
    Dim clienterec As String
    Dim s, desdes, hastas   As Integer
    desdes = 0
    hastas = 1
    
    Dim intMinRitardoInizioViaggio As Integer
'    Dim strSQL_UPDATE_ORA_PREVISTA_CONSEGNE As String

    Dim varBufDuoble As Variant

    Dim varNRSEQ As Variant

    objLogger.logDebug "Inizio ScanEvents_Palm."



    If Not DS.DA.QueryCommand(SQL_GET_EVENTS_PALM, varKBuffer, lngHowMany) Then
        objLogger.logError "Query SQL_GET_EVENTS_PALM."
        GoTo ErrHandler
    End If
    
    Do While lngHowMany = 1
        
        IdRegistro = varKBuffer(0, 0)
        
           Set VVT = New CViaggio
           With VVT
               Set .DS.DA = DS.DA
               Set .DS.ErrH = DS.DA.ErrH
           End With
        
           Set VVP = New CViaggioProdotto
           With VVP
               Set .DS.DA = DS.DA
               Set .DS.ErrH = DS.DA.ErrH
           End With
        
           Set VVC = New CViaggioConsegna
           With VVC
               Set .DS.DA = DS.DA
               Set .DS.ErrH = DS.DA.ErrH
           End With
        
           Set AAU = New CAutista
           With AAU
               Set .DS.DA = DS.DA
               Set .DS.ErrH = DS.DA.ErrH
           End With
        '
        '    objLogger.logInfo "Eventi trovati: " & lngHowMany
        '
           If lngHowMany > 0 Then
        '        For i = LBound(varKBuffer, 2) To UBound(varKBuffer, 2)
        '            ' Inizia transazione
        '            DS.DA.BeginTransaction
        '
                   varIDViaggio = ""
                   varIDConsegna = ""
                   varIDProdotto = ""
                   strTipoEvento = vbNullString
        '
        '            objLogger.logInfo "Evento n. " & i + 1
        '            objLogger.logInfo "-------------------------------------------------"
        '            objLogger.logInfo "IDSMS         [" & varKBuffer(0, i) & "]"
        '            objLogger.logInfo "NRTEL         [" & varKBuffer(1, i) & "]"
        '            objLogger.logInfo "DATASMS       [" & varKBuffer(2, i) & "]"
        '            objLogger.logInfo "DATALOC       [" & varKBuffer(3, i) & "]"
        '            objLogger.logInfo "TESTO         [" & varKBuffer(4, i) & "]"
        '            objLogger.logInfo "-------------------------------------------------"
        '
                   varDataEvento = Format(varKBuffer(8, 0), "dd/mm/yyyy")
                   varOraEvento = Format(varKBuffer(8, 0), "hh:nn")
                   varIDViaggio = varKBuffer(6, 0)
                   eventooesia = varKBuffer(5, 0)
                   varIdevento = varKBuffer(4, 0)
                    
                   If varIdevento = "005" Then
                   
                       If Not (DS.DA.QueryCommand(SQL_PALM_GET_CUANTOSCLI, varkbuffercliord, lngHowManyordcli, varIDViaggio, eventooesia)) Then
                               objLogger.logError "Query SQL_PALM_GET_ORDCLI."
                       GoTo ErrHandler
                       End If
            
                       clienterec = varkbuffercliord(0, 0)
        
                       If Not (DS.DA.QueryCommand(SQL_PALM_GET_ORDENESCLI, varkbufferordcli, lngHowManyordcli2, varIDViaggio, clienterec)) Then
                               objLogger.logError "Query SQL_PALM_GET_ORDCLI."
                       GoTo ErrHandler
                       End If
                       desdes = LBound(varkbufferordcli, 2)
                       hastas = UBound(varkbufferordcli, 2)
                   End If
                   
                   

                   For s = desdes To hastas

                   If varIdevento = "005" Then eventooesia = varkbufferordcli(0, s)

                   lngHowMany = 0
                   If Not (DS.DA.QueryCommand(SQL_PALM_GET_ORD, varkbufferord, lngHowMany, varIDViaggio)) Then
                           objLogger.logError "Query SQL_PALM_GET_ORD."
                   GoTo ErrHandler
        
                   End If
                   For i = LBound(varkbufferord, 2) To UBound(varkbufferord, 2)
                   ' Determiniamo se si tratta di un evento di fine carico o di inizio scarico
                   Select Case varIdevento
                       Case "005"
                       
                           ' Non è un carico, cerchiamo tra gli scarichi
                                If Not (DS.DA.QueryCommand(SQL_PALM_GET_CONSEGNA, varBufferConsegna, lnghowmanycons, eventooesia, varIDViaggio)) Then
                                        objLogger.logError "Query SQL_PALM_GET_CONSEGNA."
                                GoTo ErrHandler
                               End If
                               varIDConsegna = varBufferConsegna(0, 0)
                               strTipoEvento = "INI"
                           
                       Case "003"
                       
                       ' Si tratta di un carico
                           varIDConsegna = vbNullString
                           varIDProdotto = vbNullString
                           strTipoEvento = "SAL"
                    'GoTo ErrHandler
        
                       Case Else
        
                           objLogger.logError "Errore durante la ricerca del Viaggio, Autista " & varIDAutista & ", Data " & varDataEvento
                           DS.DA.EndTrans etRollback
                           'GoTo NextIteration
              
                   End Select
                   
                   'MsgBox (varIDViaggio & " " & varIdevento & " " & varOraEvento & " " & varDataEvento & " " & varIDConsegna)
        
                   objLogger.logInfo "RESOil Keys VVTIDVVT = """ & varIDViaggio & _
                                       """, VVCIDVVC = """ & varIDConsegna & """"
        
                   ' Se non trovo il viaggio, skip
                   If varIDViaggio <> "" Then
                       ' Creazione dell'evento
                       DS.Fields("0").NewValue = DS.DA.GetCounter("anaganu0", "anuprogr", "anuidanu", "904", 0)
                       DS.Fields("0").Updated = True
        
                       DS.Fields("1").NewValue = varIDViaggio
                       DS.Fields("1").Updated = True
        
                       DS.Fields("2").NewValue = varIDConsegna
                       DS.Fields("2").Updated = True
        
                       DS.Fields("3").NewValue = varIDProdotto
                       DS.Fields("3").Updated = True
        
                       DS.Fields("5").NewValue = varDataEvento
                       DS.Fields("5").Updated = True
        
                       DS.Fields("9").NewValue = "K"
                       DS.Fields("9").Updated = True
        
                       ' Carica il viaggio
                       With VVT
                           .DS.Fields.ResetFields
                           .DS.Fields("0").FindValue = varIDViaggio
        
                           If Not .Find Then
                               objLogger.logError "Find Viaggio."
                               DS.DA.EndTrans etRollback
                               GoTo NextIteration
                               'GoTo ErrHandler
                           End If
        
                           blnUsaDatiGPS = True
                       End With
        
        
                       Select Case strTipoEvento
                           ' -------------------------------------------------------
                           ' Evento - Salida de factora (Fine del carico)
                           ' -------------------------------------------------------
                           Case "SAL"
                               DS.Fields("4").NewValue = "CF"
                               DS.Fields("4").Updated = True
        
                               objLogger.logInfo "Tipo evento: Fine carico"
                               If Not Insert Then
                                   objLogger.logError "Insert ResOil evento."
                                   DS.DA.EndTrans etRollback
                                   GoTo NextIteration
                                   'GoTo ErrHandler
                               End If
        
                               ' Aggiorna lo stato del viaggio (VVTFLSTA)
                               VVT.DS.Fields("48").NewValue = VIAGGIO_STATO_INCORSO
                               VVT.DS.Fields("48").Updated = True
                               LogFieldValue VVT.DS.Fields("48")
        
                               ' Aggiorna il tempo di carico sul viaggio (VVTTMPCA)
                               If blnUsaDatiGPS Then
                                   ' -- DP 06/05/08
                                   ' -- Oltre alle ore aggiorno anche le date visto che alcuni carichi avvengono a cavallo della mezzanotte
        
                                   ' Aggiorna data e ora di inizio viaggio e data ed ora di inizio carico sulla testata viaggio (VVTDTINV, VVTHRINV, VVTDTINC, VVTHRINC)
                                   VVT.DS.Fields("6").NewValue = varDataEvento
                                   VVT.DS.Fields("6").Updated = True
                                   LogFieldValue VVT.DS.Fields("6")
                                   VVT.DS.Fields("7").NewValue = varOraEvento
                                   VVT.DS.Fields("7").Updated = True
                                   LogFieldValue VVT.DS.Fields("7")
                                   VVT.DS.Fields("4").NewValue = varDataEvento
                                   VVT.DS.Fields("4").Updated = True
                                   LogFieldValue VVT.DS.Fields("4")
                                   VVT.DS.Fields("5").NewValue = varOraEvento
                                   VVT.DS.Fields("5").Updated = True
                                   LogFieldValue VVT.DS.Fields("5")
        
                                   ' Aggiorna flag carico completato
                                   VVT.DS.Fields("39").NewValue = "S"
                                   VVT.DS.Fields("39").Updated = True
                                   LogFieldValue VVT.DS.Fields("39")
        
                                   '-- DP 05/06/08
                                   '-- Aggiorno il campo VVTFLAGG, che contiene l'indicazione dell'ultimo processo che ha aggiornato il record
                                   VVT.DS.Fields("52").NewValue = "ME"
                                   VVT.DS.Fields("52").Updated = True
                                   LogFieldValue VVT.DS.Fields("52")
        
        
                                   End If
        
        
                                    ' Aggiorna il flag modificato (VVPFLST5) di tutti i prodotti della consegna
                                    If Not DS.DA.ExecuteCommand(SQL_UPDATE_VIAGGIO_MODIFICATO, varIDViaggio) Then
                                        objLogger.logError "Errore in aggiornamento flag Modificato del viaggio."
                                        DS.DA.EndTrans etRollback
                                        GoTo NextIteration
                                        'GoTo ErrHandler
                                    End If

                               ' Aggiorna il viaggio
                               With VVT
                                   If .CheckViaggioCompletato(varIDViaggio) Then
                                       If Not .SetStato("C") Then
                                           objLogger.logError "Set Viaggio Completato."
                                           DS.DA.EndTrans etRollback
                                           GoTo NextIteration
                                           'GoTo ErrHandler
                                       End If
                                   End If
        
                                   If blnUsaDatiGPS Then
                                       ' Aggiorna il flag trasmissione viaggio
                                       .DS.Fields("41").NewValue = "S"
                                       .DS.Fields("41").Updated = True
                                       LogFieldValue .DS.Fields("41")
                                   End If
        
                                   ' Salva
                                   If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
                                       objLogger.logError "Save Viaggio."
                                       DS.DA.EndTrans etRollback
                                       GoTo NextIteration
                                       'GoTo ErrHandler
                                   End If
        
                                   .DS.DA.EndTrans etCommit
                               
                               End With
                               
                               
                           ' -------------------------------------------------------
                           ' Evento - Inicio descarga (Inizio scarico prodotto)
                           ' -------------------------------------------------------
                           Case "INI"
                               '-- DP 24/07/08
                               '-- Se c'è almeno un prodotto già trasmesso su GES non devo più aggiornare i dati
                               If Not VVP.blnGESAlreadyTransmitted(CStr(varIDViaggio), CStr(varIDConsegna)) Then
                                   objLogger.logInfo "Tipo evento: Inizio consegna prodotto"
        
                                   ' Carica la consegna e aggiorna i dati
                                   With VVC
                                       .DS.Fields.ResetFields
                                       .DS.Fields("0").FindValue = varIDViaggio
                                       .DS.Fields("1").FindValue = varIDConsegna
        
                                       If Not .Find Then
                                           objLogger.logError "Find Consegna."
                                           DS.DA.EndTrans etRollback
                                           GoTo NextIteration
                                           'GoTo ErrHandler
                                       End If
        
                                       varNRSEQ = .DS.Fields("2").Value
        
                                       If blnUsaDatiGPS Then
                                           '-- DP 15/05/08
                                           '-- Aggiorno data e ora di arrivo sui nuovi campi VVCDTARR e VVCHRARR anziché VVCTEMVG
                                           .DS.Fields("49").NewValue = varDataEvento
                                           .DS.Fields("49").Updated = True
                                           LogFieldValue .DS.Fields("49")
                                           .DS.Fields("47").NewValue = varOraEvento
                                           .DS.Fields("47").Updated = True
                                           LogFieldValue .DS.Fields("47")
        
                                           ' Aggiorna il flag consegna ultimata
                                           .DS.Fields("36").NewValue = "S"
                                           .DS.Fields("36").Updated = True
                                           LogFieldValue .DS.Fields("36")
        
                                           ' Aggiorna lo stato 'Completo'
                                           .DS.Fields("44").NewValue = "2"
                                           .DS.Fields("44").Updated = True
                                           LogFieldValue .DS.Fields("44")
        
                                           '-- DP 05/06/08
                                           '-- Aggiorno il campo VVCFLAGG, che contiene l'indicazione dell'ultimo processo che ha aggiornato il record
                                           .DS.Fields("50").NewValue = "ME"
                                           .DS.Fields("50").Updated = True
                                           LogFieldValue .DS.Fields("50")
                                       End If
        
                                       ' Salva le modifiche sulla consegna
                                       If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
                                           objLogger.logError "Save Consegna."
                                           DS.DA.EndTrans etRollback
                                           GoTo NextIteration
                                           'GoTo ErrHandler
                                       End If
                                   End With
        
        
        
                                   ' Ricerco tutti i prodotti con VVPFLSTA  <> '2'
                                   lngHowMany = 0
                                   If Not DS.DA.QueryCommand(SQL_KNOSOS_GET_COUNT_SCARICHI_NONCONCLUSI_VVT, varBuffer, lngHowMany, varIDViaggio) Then
                                       objLogger.logError "Query SQL_KNOSOS_GET_COUNT_SCARICHI_NONCONCLUSI_VVT."
                                       DS.DA.EndTrans etRollback
                                       GoTo NextIteration
                                       'GoTo ErrHandler
                                   End If
        
                                   VVP.DS.Field("Numero Viaggio").FindValue = varIDViaggio
                                   VVP.DS.Field("Numero Consegna").FindValue = varIDConsegna
                                   If Not VVP.GetList(True) Then
                                       DS.DA.EndTrans etRollback
                                       GoTo NextIteration
                                       'GoTo ErrHandler
                                   End If
        
                                   If varBuffer(0, 0) = 0 Then
                                       ' Tutti i prodotti del viaggio hanno VVPFLSTA  = '2', impostiamo il flag modificato (VVPFLST5) di tutti i prodotti del viaggio
                                       For j = 0 To VVP.DS.HowMany - 1
                                           ' Carica i prodotti della consegna e aggiorna i dati
                                           With VVP
                                               .DS.Fields.ResetFields
                                               .DS.Fields("0").FindValue = varIDViaggio
                                               .DS.Fields("1").FindValue = varIDConsegna
                                               .DS.Fields("2").FindValue = varIDProdotto
        
                                               If Not .Find Then
                                                   objLogger.logError "Find Consegna Prodotto."
                                                   DS.DA.EndTrans etRollback
                                                   GoTo NextIteration
                                                   'GoTo ErrHandler
                                               End If
                                               If .DS.HowMany = 0 Then
                                                   objLogger.logError "Prodotto Consegna non trovato"
                                                   DS.DA.EndTrans etRollback
                                                   GoTo NextIteration
                                                   'GoTo ErrHandler
                                               End If
                                           End With
                                       Next j
        
                                       If Not DS.DA.ExecuteCommand(SQL_UPDATE_VIAGGIO_MODIFICATO, varIDViaggio) Then
                                           objLogger.logError "Errore in aggiornamento flag Modificato del viaggio."
                                           DS.DA.EndTrans etRollback
                                           GoTo NextIteration
                                           'GoTo ErrHandler
                                       End If
                                   Else
                                       For j = 0 To VVP.DS.HowMany - 1
                                           varIDProdotto = VVP.DS.ListItem(0, j)
        
                                           DS.Fields("0").NewValue = DS.DA.GetCounter("anaganu0", "anuprogr", "anuidanu", "904", 0)
                                           DS.Fields("0").Updated = True
        
                                           DS.Fields("1").NewValue = varIDViaggio
                                           DS.Fields("1").Updated = True
        
                                           DS.Fields("2").NewValue = varIDConsegna
                                           DS.Fields("2").Updated = True
        
                                           DS.Fields("3").NewValue = varIDProdotto
                                           DS.Fields("3").Updated = True
        
                                           DS.Fields("4").NewValue = "SI"
                                           DS.Fields("4").Updated = True
        
                                           DS.Fields("5").NewValue = varDataEvento
                                           DS.Fields("5").Updated = True
        
                                           DS.Fields("9").NewValue = "K"
                                           DS.Fields("9").Updated = True
        
                                           objLogger.logInfo "Creazione dell'evento in ResOil."
                                           If Not Insert Then
                                               objLogger.logError "Insert ResOil evento."
                                               DS.DA.EndTrans etRollback
                                               GoTo NextIteration
                                               'GoTo ErrHandler
                                           End If
        
                                           ' Carica i prodotti della consegna e aggiorna i dati
                                           With VVP
                                               .DS.Fields.ResetFields
                                               .DS.Fields("0").FindValue = varIDViaggio
                                               .DS.Fields("1").FindValue = varIDConsegna
                                               .DS.Fields("2").FindValue = varIDProdotto
        
                                               If Not .Find Then
                                                   objLogger.logError "Find Consegna Prodotto."
                                                   DS.DA.EndTrans etRollback
                                                   GoTo NextIteration
                                                   'GoTo ErrHandler
                                               End If
                                               If .DS.HowMany = 0 Then
                                                   objLogger.logError "Prodotto Consegna non trovato"
                                                   DS.DA.EndTrans etRollback
                                                   GoTo NextIteration
                                                   'GoTo ErrHandler
                                               End If
        
                                               ' Aggiorna lo stato dello scarico (VVPFLSTA)
                                               .DS.Fields("37").NewValue = "2"
                                               .DS.Fields("37").Updated = True
                                               LogFieldValue .DS.Fields("37")
        
                                               ' Aggiorna il flag Modificato (VVPFLST5)
                                               .DS.Field("Stato Riga 5").NewValue = "S"
                                               .DS.Field("Stato Riga 5").Updated = True
                                               LogFieldValue .DS.Field("Stato Riga 5")
        
                                               '-- DP 05/06/08
                                               '-- Aggiorno il campo VVPFLAGG, che contiene l'indicazione dell'ultimo processo che ha aggiornato il record
                                               .DS.Fields("43").NewValue = "ME"
                                               .DS.Fields("43").Updated = True
                                               LogFieldValue .DS.Fields("43")
        
                                               ' Aggiorna l'ora inizio scarico (VVPHRINC)
                                               If blnUsaDatiGPS Then
                                                   .DS.Fields("38").NewValue = varOraEvento
                                                   .DS.Fields("38").Updated = True
                                                   LogFieldValue .DS.Fields("38")
                                               End If
        
                                               ' Salva
                                               If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
                                                   objLogger.logError "Save Consegna Prodotto."
                                                   DS.DA.EndTrans etRollback
                                                   GoTo NextIteration
                                                   'GoTo ErrHandler
                                               End If
                                           End With
                                       Next j
                                   End If
        
                                   ' Aggiorna il viaggio
                                   With VVT
                                       If .CheckViaggioCompletato(varIDViaggio) Then
                                           If Not .SetStato("C") Then
                                               objLogger.logError "Set Viaggio Completato."
                                               DS.DA.EndTrans etRollback
                                               GoTo NextIteration
                                               'GoTo ErrHandler
                                           End If
                                       End If
        
                                       If blnUsaDatiGPS Then
                                           ' Aggiorna il flag trasmissione viaggio
                                           .DS.Fields("41").NewValue = "S"
                                           .DS.Fields("41").Updated = True
                                           LogFieldValue .DS.Fields("41")
        
                                           '-- DP 05/06/08
                                           '-- Aggiorno il campo VVTFLAGG, che contiene l'indicazione dell'ultimo processo che ha aggiornato il record
                                           .DS.Fields("52").NewValue = "ME"
                                           .DS.Fields("52").Updated = True
                                           LogFieldValue .DS.Fields("52")
                                       End If
        
                                       ' Salva
                                       If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
                                           objLogger.logError "Save Viaggio."
                                           DS.DA.EndTrans etRollback
                                           GoTo NextIteration
                                           'GoTo ErrHandler
                                       End If
        
                                       .DS.DA.EndTrans etCommit
                                   End With
                               Else
                                   objLogger.logError "Ora di reale di consegna già trasmessa su GES"
                               End If
        
                               DS.DA.EndTrans etCommit
        
                           Case Else
                               objLogger.logError "Tipo evento sconosciuto."
                               DS.DA.EndTrans etRollback
                               GoTo NextIteration
                               'GoTo ErrHandler
        
                       End Select
                   Else
                       DS.DA.LogExecTime "WARN: Viaggio non trovato!", -1
                   End If
        
                   DS.DA.EndTrans etCommit
NextIteration:
               Next i
               Next s
           End If
        
        'Dim evoes As New CDataAccess
        '
        '    If Not (evoes.ExecuteQuery("UPDATE EventosMBFlotasmanual Set Capturado = 1 WHERE (IdRegistro = '" & Trim(idregistro) & "')")) Then
        '    End If
           Dim SQL_UPDATE_MBFlotas As String
           SQL_UPDATE_MBFlotas = "UPDATE EventosMBFlotasmanual Set Capturado = 1 WHERE (IdRegistro = ?)"
           
           If Not DS.DA.ExecuteCommand(SQL_UPDATE_MBFlotas, IdRegistro) Then
           End If
            If Not DS.DA.QueryCommand(SQL_GET_EVENTS_PALM, varKBuffer, lngHowMany) Then
            objLogger.logError "Query SQL_GET_EVENTS_PALM."
            GoTo ErrHandler
            End If

    Loop
    ScanEvents_Palm = True

    Set VVP = Nothing
    Set VVT = Nothing
    Set VVC = Nothing

    objLogger.logDebug "Fine ScanEvents_palm senza errori."
    GoTo EndProc

ErrHandler:
    DS.DA.EndTrans etRollback

    Set VVP = Nothing
    Set VVT = Nothing
    Set VVC = Nothing

    objLogger.logError "Fine ScanEvents_SMS con ERRORI."

    ScanEvents_Palm = False

EndProc:


End Function
Function ScanEvents_Cisternas() As Boolean
'
'    'hay que añadir el codigo del vehiculo para hacer la consulta rapida.
'    If DS.DA.QueryCommand("Select * From EventosCisternasMBFlotas where capturado = 0 ", varKBuffer, lngHowMany) Then
'    End If
'    'tenemos los eventos no capturados de las cisternas, buscamos el evento 004 que esta parado y rellenamos los compartimentos
'    'Cuales son los parsametros de este campo? TR,SP,DC,DP,XX
'    If DS.DA.QueryCommand("Select * From VIAGGES0 WHERE IDVEHI= " + IDV + " AND GESCODEV = '004' ", varKBuffer, lngHowMany) Then
'    End If
'    'una vez tenemos el evento 004 nos guardamos el registro y comprobamos que no existan eventos 005 del mismo viaje
'
'    If DS.DA.QueryCommand("Select * From VIAGGES0 WHERE IDVEHI= " + IDV + " AND GESCODEV = '005' ", varKBuffer, lngHowMany) Then
'    End If
'
'    If Not (lngHowMany > 0) Then
'        'rellenamos todos los eventos 004 del viaje que no tengan un evento 005
'        'para saber que comartimento tenemos que rellenar selecionamos por deposito
'        '01= GESCOMP1
'        '02= GESCOMP2
'        '03= GESCOMP3
'        '04= GESCOMP4
'        '05= GESCOMP5
'        '06= GESCOMP6
'        '07= GESCOMP7
'
'        ' que pasa con los 004 a medias que no llegan a tiempo antes que el 005? y los que no llega naDA DE NADA?
'
'
'        ' si el compartimento es el 7 es el ultimo... marcar GESTRCEP = 1
'    End If
'
'    For i = 0 To i = UBound(varKBuffer, 2)
'        If DS.DA.ExecuteQuery("UPDATE EventosCisternasMBFlotas SET Capturado = '1' WHERE IDRegistro= " + IDR + " ") Then
'        End If
'    Next i
    
End Function
Function ScanEvents_Palm_nuevo() As Boolean
    Dim i As Long
    Dim j As Long
    Dim varKBuffer As Variant
    Dim varBuffer As Variant
    Dim varKBuffer2 As Variant
    Dim varkbufferord As Variant
    Dim varkbuffercliord As Variant
    Dim varkbufferordcli As Variant
    Dim varBufferConsegna As Variant
    Dim varDataEvento As Variant
    Dim varOraEvento As Variant
    Dim varTelAutista As Variant
    Dim strLastDate As String
    Dim lngHowMany As Long
    Dim lngHowManyordcli As Long
    Dim lngHowManyordcli2 As Long
    Dim lngcomplet As Long
    Dim lnghowmanycons As Long
    Dim rowcep As Variant
    Dim lngcep As Long
    Dim varIDViaggio As String
    Dim varIDConsegna, Rowcomplet As Variant
    Dim varIDProdotto As Variant
    Dim varIDAutista As Variant
    Dim varIDSMS As Variant
    Dim varTestoSMS As Variant
    Dim VVP As CViaggioProdotto
    Dim VVT As CViaggio
    Dim VVC As CViaggioConsegna
    Dim AAU As CAutista
    Dim dateLocalEventDateTime As Date
    Dim sngKm As Single
    Dim strSQLStatement As String
    Dim strTipoEvento As String
    Dim blnUsaDatiGPS As Boolean
    Dim varIdevento As String
    Dim IdRegistro As String
    Dim eventooesia As String
    Dim clienterec As String
    Dim s, desdes, hastas, desdeord, hastaord As Integer
    Dim esdoble As Boolean
    esdoble = False
    desdes = 0
    hastas = 0
    desdeord = 0
    hastaord = 0
    Dim restr As Variant
    Dim HREv As String
    Dim DYEv As String
    Dim CordX As String
    Dim CordY As String
    Dim IDPalm As String
    Dim Estadoviaj As Variant
    Dim lnghm As Long
    Dim Orden As String
    Dim IdTarea As String
    Dim Gen_Por As String
    Dim ProdObserva As String
   ' Dim IdRegistro As String
    
    Dim intMinRitardoInizioViaggio As Integer
'    Dim strSQL_UPDATE_ORA_PREVISTA_CONSEGNE As String

    Dim varBufDuoble As Variant

    Dim varNRSEQ As Variant

    objLogger.logDebug "Inizio ScanEvents_Palm."
    
    
    ' Jg 31/1/18  Gamusino los eventos 003 no se pegan a veces, ponemos el contador en entaenloc en Evwentosmbflotas manual
    
    Dim SQL_EVENTOSERR As String
    Dim SQLUPDATEEVENTOERR As String
    Dim Eventoerr As String
    Dim Intentos As Integer
    Dim varKbuffererr As Variant
    Dim lngHowmerr As Long
    
    SQL_EVENTOSERR = "SELECT        EventosMBFlotasmanual.IdRegistro, EventosMBFlotasmanual.IdUsuarioMB, EventosMBFlotasmanual.CodigoEvento, EventosMBFlotasmanual.CodigoServicio, EventosMBFlotasmanual.FechaEvento, " & _
                        "                         EventosMBFlotasmanual.FechaServidor, EventosMBFlotasmanual.Capturado, VIAGVVT0.VVTFLST3, VIAGVVT0.VVTIDASC, EventosMBFlotasmanual.EntaEnLoc " & _
                        "FROM            EventosMBFlotasmanual INNER JOIN " & _
                        "                         VIAGVVT0 ON EventosMBFlotasmanual.CodigoServicio COLLATE SQL_Latin1_General_CP1_CI_AS = VIAGVVT0.VVTIDVVT " & _
                        "WHERE        (EventosMBFlotasmanual.CodigoEvento = '003') AND (EventosMBFlotasmanual.FechaServidor > GETDATE() - 3) AND (VIAGVVT0.VVTFLST3 IS NULL)" & _
                        "ORDER BY EventosMBFlotasmanual.IdRegistro DESC "
    
    If Not DS.DA.QueryCommand(SQL_EVENTOSERR, varKbuffererr, lngHowmerr) Then
         objLogger.logError "Query SQL_GET_EVENTS_PALM."
         GoTo ErrHandler
     End If
    
    
    ' Si hay errores entramos
    If lngHowmerr > 0 Then
    
        For i = LBound(varKbuffererr, 2) To UBound(varKbuffererr, 2)
        
            Eventoerr = varKbuffererr(0, i)
            Intentos = varKbuffererr(9, i)
            
            Intentos = Intentos + 1
            
            SQLUPDATEEVENTOERR = "UPDATE       EventosMBFlotasmanual SET Capturado = 0, EntaEnLoc = '" & Intentos & "'WHERE(IdRegistro = '" & Eventoerr & "')"
            If Not DS.DA.ExecuteQuery(SQLUPDATEEVENTOERR) Then
            End If

        
        Next i
    
    End If
    
    



    If Not DS.DA.QueryCommand(SQL_GET_EVENTOS_PALM_TOTAL, varKBuffer, lngHowMany) Then
        objLogger.logError "Query SQL_GET_EVENTS_PALM."
        GoTo ErrHandler
    End If
    
    

    Do While lngHowMany = 1
    
    HREv = CStr(varKBuffer(17, 0))
    DYEv = CStr(varKBuffer(16, 0))
    CordY = varKBuffer(12, 0)
    CordX = varKBuffer(13, 0)
    IDPalm = varKBuffer(2, 0)
    IdRegistro = varKBuffer(0, 0)
    ' JG 25/11/21 Campo observaciones incidencias utilizamos el campo productos ya existente en eventosmanuales
    If Not IsNull(varKBuffer(7, 0)) Then
            ProdObserva = varKBuffer(7, 0)
            Else
            ProdObserva = ""
            End If
    
        IdRegistro = varKBuffer(0, 0)
        
           Set VVT = New CViaggio
           With VVT
               Set .DS.DA = DS.DA
               Set .DS.ErrH = DS.DA.ErrH
           End With
        
           Set VVP = New CViaggioProdotto
           With VVP
               Set .DS.DA = DS.DA
               Set .DS.ErrH = DS.DA.ErrH
           End With
        
           Set VVC = New CViaggioConsegna
           With VVC
               Set .DS.DA = DS.DA
               Set .DS.ErrH = DS.DA.ErrH
           End With
        
           Set AAU = New CAutista
           With AAU
               Set .DS.DA = DS.DA
               Set .DS.ErrH = DS.DA.ErrH
           End With
     
           If lngHowMany > 0 Then
       
                   varIDViaggio = ""
                   varIDConsegna = ""
                   varIDProdotto = ""
                   strTipoEvento = vbNullString
        
     
                   varDataEvento = Format(varKBuffer(8, 0), "dd/mm/yyyy")
                   varOraEvento = Format(varKBuffer(8, 0), "hh:nn")
                   If Not IsNull(varKBuffer(6, 0)) Then
                        varIDViaggio = CStr(varKBuffer(6, 0))
                        Else
                        varIDViaggio = ""
                        End If
                   varIdevento = varKBuffer(4, 0)
                    
                   If varIdevento = "005" Or varIdevento = "006" Or varIdevento = "999" Or (CInt(varIdevento) > 7 And CInt(varIdevento) < 26) Or (CInt(varIdevento) = 34) Then
                   eventooesia = varKBuffer(5, 0)
                   If Not (DS.DA.QueryCommand(SQL_PALM_GET_CUANTOSCLI, varkbuffercliord, lngHowManyordcli, varIDViaggio, eventooesia)) Then
                           objLogger.logError "Query SQL_PALM_GET_CUANTOSCLI."
                   GoTo ErrHandler
                   End If
                   
                   If lngHowManyordcli > 0 Then
                        If (IsNull(varkbuffercliord(0, 0)) Or IsEmpty(varkbuffercliord(0, 0))) Then
                        clienterec = ""
                        Else
                        clienterec = varkbuffercliord(0, 0)
                        End If
                                                                
        
                       If Not (DS.DA.QueryCommand(SQL_PALM_GET_ORDENESCLI, varkbufferordcli, lngHowManyordcli2, varIDViaggio, clienterec)) Then
                               objLogger.logError "Query SQL_PALM_GET_ORDCLI."
                       GoTo ErrHandler
                       End If
                       desdes = LBound(varkbufferordcli, 2)
                       hastas = UBound(varkbufferordcli, 2)
                    Else
                     GoTo NextIteration
                    End If
                    
                   End If
                   
                   

  For s = desdes To hastas
esdoble = False
                   If varIdevento = "005" Then
                      eventooesia = varkbufferordcli(0, s)
                   End If
                   

                 lngHowMany = 0
                 If Not (DS.DA.QueryCommand(SQL_PALM_GET_ORD, varkbufferord, lngHowMany, varIDViaggio)) Then
                    objLogger.logError "Query SQL_PALM_GET_ORD."
                 End If
                 If (lngHowMany > 1) Then
                   esdoble = True
                 End If
                 If (lngHowMany > 0) Then
                    desdeord = LBound(varkbufferord, 2)
                    hastaord = UBound(varkbufferord, 2)
                 End If
                 
                   
i = 0
 '   For i = desdeord To hastaord

                   ' Determiniamo se si tratta di un evento di fine carico o di inizio scarico
                      Select Case varIdevento
                   
                       Case "003"
                            ' Si tratta di un carico
                                varIDConsegna = vbNullString
                                varIDProdotto = vbNullString
                                strTipoEvento = "SAL"
                                 
                       Case "005"
                               ' Non è un carico, cerchiamo tra gli scarichi
                               If Not (DS.DA.QueryCommand(SQL_PALM_GET_CONSEGNA, varBufferConsegna, lnghowmanycons, eventooesia, varIDViaggio)) Then
                                        objLogger.logError "Query SQL_PALM_GET_CONSEGNA."
                               GoTo ErrHandler
                               End If
                               varIDConsegna = varBufferConsegna(0, 0)
                               strTipoEvento = "INI"
        
                       Case Else
        
                           objLogger.logError "Errore durante la ricerca del Viaggio, Autista " & varIDAutista & ", Data " & varDataEvento
                           DS.DA.EndTrans etRollback
                           'GoTo NextIteration
              
                   End Select
                   
 ' Asignamos al campo Gen_por quien ha generado el evento M-Manual R-Resoil S-SMS P-PDA
                   
    Select Case varKBuffer(2, 0)

        Case "0"
        Gen_Por = "M"
        
        Case "999"
        Gen_Por = "R"

        Case "666"
        Gen_Por = "S"
        
        Case Else
        Gen_Por = "P"
        
    End Select

 'SM Omitir el viaje si este esta completado.
   If Gen_Por <> "M" And Gen_Por <> "R" Then
        If (varIdevento = "003") Then
            If DS.DA.QueryCommand("SELECT VVTFLST3 FROM VIAGVVT0 WHERE VVTIDVVT= ?", Estadoviaj, lnghm, varIDViaggio) Then
            End If
            If (lnghm > 0) Then
                If (Estadoviaj(0, 0) = "S") Then
                    GoTo NextIteration
                End If
            End If
        End If
        If (varIdevento = "005") Then
            If DS.DA.QueryCommand("SELECT  VVCFLST5 From VIAGVVC0 WHERE     (VVCIDVVT = ?) AND (VVCIDOCT = ?)", Estadoviaj, lnghm, varIDViaggio, eventooesia) Then
            End If
            If (lnghm > 0) Then
                If Not (Estadoviaj(0, 0) <> "S") Then
                    GoTo NextIteration
                End If
            End If
        End If
        If (CInt(varIdevento) > 8 And CInt(varIdevento) < 18) Then
              GoTo NextIteration
        End If
   End If

'Comprobamos el operador del viaje
If Not (DS.DA.QueryCommand("SELECT VVTIDASC From VIAGVVT0 WHERE (VVTIDVVT = ?)", rowcep, lngcep, varIDViaggio)) Then
End If
Dim contcep As String

If (lngcep > 0) Then
   contcep = rowcep(0, 0)
Else
   If varIdevento = "001" Or varIdevento = "007" Then
   contcep = "038"
   Else
   contcep = ""
   End If
End If

     'comprobamos si el evento esta dentro de horario, si esta 15 min +/- del horario lo metemos en hora.
    If (varIdevento = "005") And (contcep = "038") Then
                Dim desde1 As Date
                Dim desde2 As Date
                Dim desde3 As Date
                Dim hasta1 As Date
                Dim hasta2 As Date
                Dim hasta3 As Date
                Dim margen As Date
                Dim horafinal As Date
                Dim VarKbufferHC As Variant
                Dim lngGesHC As Long
                
                ' Establecemos en margen el margen de horarario que nos dan para cumplir
                margen = "01:00:00"
                
                      If DS.DA.QueryCommand(SQL_GES_HorariosCli, VarKbufferHC, lngGesHC, varIDViaggio, clienterec) Then
                      If Not IsNull(VarKbufferHC(0, 0)) Then desde1 = VarKbufferHC(0, 0) Else desde1 = "00:00"
                      If Not IsNull(VarKbufferHC(1, 0)) Then hasta1 = VarKbufferHC(1, 0) Else hasta1 = "00:00"
                      If Not IsNull(VarKbufferHC(2, 0)) Then desde2 = VarKbufferHC(2, 0) Else desde2 = "00:00"
                      If Not IsNull(VarKbufferHC(3, 0)) Then hasta2 = VarKbufferHC(3, 0) Else hasta2 = "00:00"
                      If Not IsNull(VarKbufferHC(4, 0)) Then desde3 = VarKbufferHC(4, 0) Else desde3 = "00:00"
                      If Not IsNull(VarKbufferHC(5, 0)) Then hasta3 = VarKbufferHC(5, 0) Else hasta3 = "00:00"

                 
                
                If (desde1 - margen) < #1/1/1900# Or desde1 = #1/1/1900# Then
                 desde1 = #1/1/1900 12:00:01 AM#
                 Else: desde1 = desde1 - margen
                End If
                      
                If (hasta1 + margen) > #1/1/1900 11:59:00 PM# Or hasta1 = #1/1/1900# Then
                 hasta1 = #1/1/1900 11:59:00 PM#
                 Else: hasta1 = hasta1 + margen
                End If
                                                                         
                If (desde2 - margen) < #1/1/1900# Then
                 desde2 = #1/1/1900 12:00:01 AM#
                 Else: desde2 = desde2 - margen
                End If
                      
                If (hasta2 + margen) > #1/1/1900 11:59:00 PM# Then
                 hasta2 = #1/1/1900 11:59:00 PM#
                 Else: hasta2 = hasta2 + margen
                End If
                    
                If (desde3 - margen) < #1/1/1900# Then
                 desde3 = #1/1/1900 12:00:01 AM#
                 Else: desde3 = desde3 - margen
                End If
                      
                If (hasta3 + margen) > #1/1/1900 11:59:00 PM# Then
                 hasta3 = #1/1/1900 11:59:00 PM#
                 Else: hasta3 = hasta3 + margen
                End If
                      
                      
                        If (CDate("1/1/1900 " + CStr(HREv)) > desde1 And CDate("1/1/1900 " + CStr(HREv)) < hasta1) Then

                        ElseIf (CDate("1/1/1900 " + CStr(HREv)) > desde2 And CDate("1/1/1900 " + CStr(HREv)) < hasta2) Then

                        ElseIf (CDate("1/1/1900 " + CStr(HREv)) > desde3 And CDate("1/1/1900 " + CStr(HREv)) < hasta3) Then

                        Else
                        Dim HoraI, HoraF As Date
                        HoraI = CDate(HREv) + "00:15"

                        HoraI = Format(HoraI, "hh:mm")

                        HoraF = TimeValue(CDate(HREv)) - TimeValue("00:15")
                        HoraF = Format(HoraF, "hh:mm")


                            'Consultar si esta dentro de 15 min antes/despues y aumentarlo o retrasarlo 15 min
                            'Comprobar si esta dentro del campo restandole 15 min a la hora final y sumandole 15 a la inicial, si es así comprobamos que la hora sea mas grande que el horario de entrega,
                            'si es así le restaremos 15 min para que este dentro de horario, si no le sumaremos 15 min.
                                If (CDate("1/1/1900 " + CStr(HoraI)) > desde1 And CDate("1/1/1900 " + CStr(HoraF)) < hasta1) Then
                                    If (CDate("1/1/1900 " + CStr(HREv)) > hasta1) Then
                                        HREv = HoraF
                                    Else
                                        HREv = HoraI
                                    End If
                                ElseIf (CDate("1/1/1900 " + CStr(HoraI)) > desde2 And CDate("1/1/1900 " + CStr(HoraF)) < hasta2) Then
                                     If (CDate("1/1/1900 " + CStr(HREv)) > hasta2) Then
                                        HREv = HoraF
                                    Else
                                        HREv = HoraI
                                    End If
                                ElseIf (CDate("1/1/1900 " + CStr(HoraI)) > desde3 And CDate("1/1/1900 " + CStr(HoraF)) < hasta3) Then
                                     If (CDate("1/1/1900 " + CStr(HREv)) > hasta3) Then
                                        HREv = HoraF
                                    Else
                                        HREv = HoraI
                                    End If
                                End If
                        End If
                      End If

       End If
                
                   
                   
                   
                   
                   
                   
                   'MsgBox (varIDViaggio & " " & varIdevento & " " & varOraEvento & " " & varDataEvento & " " & varIDConsegna)
        
                   objLogger.logInfo "RESOil Keys VVTIDVVT = """ & varIDViaggio & _
                                       """, VVCIDVVC = """ & varIDConsegna & """"
        
                   ' Se non trovo il viaggio, skip
                   If varIDViaggio <> "" Then
                       ' Creazione dell'evento
                       DS.Fields("0").NewValue = DS.DA.GetCounter("anaganu0", "anuprogr", "anuidanu", "904", 0)
                       DS.Fields("0").Updated = True
        
                       DS.Fields("1").NewValue = varIDViaggio
                       DS.Fields("1").Updated = True
        
                       DS.Fields("2").NewValue = varIDConsegna
                       DS.Fields("2").Updated = True
        
                       DS.Fields("3").NewValue = varIDProdotto
                       DS.Fields("3").Updated = True
        
                       DS.Fields("5").NewValue = varDataEvento
                       DS.Fields("5").Updated = True
        
                       DS.Fields("9").NewValue = "K"
                       DS.Fields("9").Updated = True
        
                       ' Carica il viaggio
                       With VVT
                           .DS.Fields.ResetFields
                           .DS.Fields("0").FindValue = varIDViaggio
        
                           If Not .Find Then
                               objLogger.logError "Find Viaggio."
                               DS.DA.EndTrans etRollback
                               GoTo NextIteration
                               'GoTo ErrHandler
                           End If
        
                           blnUsaDatiGPS = True
                       End With
        
        
                       Select Case strTipoEvento
                           ' -------------------------------------------------------
                           ' Evento - Salida de factora (Fine del carico)
                           ' -------------------------------------------------------
                           Case "SAL"
                               DS.Fields("4").NewValue = "CF"
                               DS.Fields("4").Updated = True
        
                               objLogger.logInfo "Tipo evento: Fine carico"
                               If Not Insert Then
                                   objLogger.logError "Insert ResOil evento."
                                   DS.DA.EndTrans etRollback
                                   GoTo NextIteration
                                   'GoTo ErrHandler
                               End If
        
                               ' Aggiorna lo stato del viaggio (VVTFLSTA)
                               VVT.DS.Fields("48").NewValue = VIAGGIO_STATO_INCORSO
                               VVT.DS.Fields("48").Updated = True
                               LogFieldValue VVT.DS.Fields("48")
        
                               ' Aggiorna il tempo di carico sul viaggio (VVTTMPCA)
                               If blnUsaDatiGPS Then
                                   ' -- DP 06/05/08
                                   ' -- Oltre alle ore aggiorno anche le date visto che alcuni carichi avvengono a cavallo della mezzanotte
        
                                   ' Aggiorna data e ora di inizio viaggio e data ed ora di inizio carico sulla testata viaggio (VVTDTINV, VVTHRINV, VVTDTINC, VVTHRINC)
                                   VVT.DS.Fields("6").NewValue = varDataEvento
                                   VVT.DS.Fields("6").Updated = True
                                   LogFieldValue VVT.DS.Fields("6")
                                   VVT.DS.Fields("7").NewValue = varOraEvento
                                   VVT.DS.Fields("7").Updated = True
                                   LogFieldValue VVT.DS.Fields("7")
                                   VVT.DS.Fields("4").NewValue = varDataEvento
                                   VVT.DS.Fields("4").Updated = True
                                   LogFieldValue VVT.DS.Fields("4")
                                   VVT.DS.Fields("5").NewValue = varOraEvento
                                   VVT.DS.Fields("5").Updated = True
                                   LogFieldValue VVT.DS.Fields("5")
        
                                   ' Aggiorna flag carico completato
                                   VVT.DS.Fields("39").NewValue = "S"
                                   VVT.DS.Fields("39").Updated = True
                                   LogFieldValue VVT.DS.Fields("39")
        
                                   '-- DP 05/06/08
                                   '-- Aggiorno il campo VVTFLAGG, che contiene l'indicazione dell'ultimo processo che ha aggiornato il record
                                   VVT.DS.Fields("52").NewValue = "ME"
                                   VVT.DS.Fields("52").Updated = True
                                   LogFieldValue VVT.DS.Fields("52")
        
        
                                   End If
        
        
                                    ' Aggiorna il flag modificato (VVPFLST5) di tutti i prodotti della consegna
                                    If Not DS.DA.ExecuteCommand(SQL_UPDATE_VIAGGIO_MODIFICATO, varIDViaggio) Then
                                        objLogger.logError "Errore in aggiornamento flag Modificato del viaggio."
                                        DS.DA.EndTrans etRollback
                                        GoTo NextIteration
                                        'GoTo ErrHandler
                                    End If

                               ' Aggiorna il viaggio
                               With VVT
                                   If .CheckViaggioCompletato(varIDViaggio) Then
                                       If Not .SetStato("C") Then
                                           objLogger.logError "Set Viaggio Completato."
                                           DS.DA.EndTrans etRollback
                                           GoTo NextIteration
                                           'GoTo ErrHandler
                                       End If
                                   End If
        
                                   If blnUsaDatiGPS Then
                                       ' Aggiorna il flag trasmissione viaggio
                                       .DS.Fields("41").NewValue = "S"
                                       .DS.Fields("41").Updated = True
                                       LogFieldValue .DS.Fields("41")
                                   End If
        
                                   ' Salva
                                   If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
                                       objLogger.logError "Save Viaggio."
                                       DS.DA.EndTrans etRollback
                                       GoTo NextIteration
                                       'GoTo ErrHandler
                                   End If
        
                                   .DS.DA.EndTrans etCommit
                               
                               End With
                               
                               
                           ' -------------------------------------------------------
                           ' Evento - Inicio descarga (Inizio scarico prodotto)
                           ' -------------------------------------------------------
                           Case "INI"
                               '-- DP 24/07/08
                               '-- Se c'è almeno un prodotto già trasmesso su GES non devo più aggiornare i dati
'                               If Not VVP.blnGESAlreadyTransmitted(CStr(varIDViaggio), CStr(varIDConsegna)) Then
'                                   objLogger.logInfo "Tipo evento: Inizio consegna prodotto"
        
                                   ' Carica la consegna e aggiorna i dati
                                   With VVC
                                       .DS.Fields.ResetFields
                                       .DS.Fields("0").FindValue = varIDViaggio
                                       .DS.Fields("1").FindValue = varIDConsegna
        
                                       If Not .Find Then
                                           objLogger.logError "Find Consegna."
                                           DS.DA.EndTrans etRollback
                                           GoTo NextIteration
                                           'GoTo ErrHandler
                                       End If
        
                                       varNRSEQ = .DS.Fields("2").Value
        
                                       If blnUsaDatiGPS Then
                                           '-- DP 15/05/08
                                           '-- Aggiorno data e ora di arrivo sui nuovi campi VVCDTARR e VVCHRARR anziché VVCTEMVG
                                           .DS.Fields("49").NewValue = varDataEvento
                                           .DS.Fields("49").Updated = True
                                           LogFieldValue .DS.Fields("49")
                                           .DS.Fields("47").NewValue = HREv
                                           .DS.Fields("47").Updated = True
                                           LogFieldValue .DS.Fields("47")
        
                                           ' Aggiorna il flag consegna ultimata
                                           .DS.Fields("36").NewValue = "S"
                                           .DS.Fields("36").Updated = True
                                           LogFieldValue .DS.Fields("36")
        
                                           ' Aggiorna lo stato 'Completo'
                                           .DS.Fields("44").NewValue = "2"
                                           .DS.Fields("44").Updated = True
                                           LogFieldValue .DS.Fields("44")
        
                                           '-- DP 05/06/08
                                           '-- Aggiorno il campo VVCFLAGG, che contiene l'indicazione dell'ultimo processo che ha aggiornato il record
                                           .DS.Fields("50").NewValue = "ME"
                                           .DS.Fields("50").Updated = True
                                           LogFieldValue .DS.Fields("50")
                                       End If
        
                                       ' Salva le modifiche sulla consegna
                                       If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
                                           objLogger.logError "Save Consegna."
                                           DS.DA.EndTrans etRollback
                                           GoTo NextIteration
                                           'GoTo ErrHandler
                                       End If
                                       DS.DA.EndTrans etCommit
                                   End With
        
        
        
                                   ' Ricerco tutti i prodotti con VVPFLSTA  <> '2'
                                   lngHowMany = 0
                                   If Not DS.DA.QueryCommand(SQL_KNOSOS_GET_COUNT_SCARICHI_NONCONCLUSI_VVT, varBuffer, lngHowMany, varIDViaggio) Then
                                       objLogger.logError "Query SQL_KNOSOS_GET_COUNT_SCARICHI_NONCONCLUSI_VVT."
                                       DS.DA.EndTrans etRollback
                                       GoTo NextIteration
                                       'GoTo ErrHandler
                                   End If
        
                                   VVP.DS.Field("Numero Viaggio").FindValue = varIDViaggio
                                   VVP.DS.Field("Numero Consegna").FindValue = varIDConsegna
                                   If Not VVP.GetList(True) Then
                                       DS.DA.EndTrans etRollback
                                       GoTo NextIteration
                                       'GoTo ErrHandler
                                   End If
        
                                   If varBuffer(0, 0) = 0 Then
                                       ' Tutti i prodotti del viaggio hanno VVPFLSTA  = '2', impostiamo il flag modificato (VVPFLST5) di tutti i prodotti del viaggio
                                       For j = 0 To VVP.DS.HowMany - 1
                                           ' Carica i prodotti della consegna e aggiorna i dati
                                           With VVP
                                               .DS.Fields.ResetFields
                                               .DS.Fields("0").FindValue = varIDViaggio
                                               .DS.Fields("1").FindValue = varIDConsegna
                                               .DS.Fields("2").FindValue = varIDProdotto
        
                                               If Not .Find Then
                                                   objLogger.logError "Find Consegna Prodotto."
                                                   DS.DA.EndTrans etRollback
                                                   GoTo NextIteration
                                                   'GoTo ErrHandler
                                               End If
                                               If .DS.HowMany = 0 Then
                                                   objLogger.logError "Prodotto Consegna non trovato"
                                                   DS.DA.EndTrans etRollback
                                                   GoTo NextIteration
                                                   'GoTo ErrHandler
                                               End If
                                           End With
                                       Next j
        
                                       If Not DS.DA.ExecuteCommand(SQL_UPDATE_VIAGGIO_MODIFICATO, varIDViaggio) Then
                                           objLogger.logError "Errore in aggiornamento flag Modificato del viaggio."
                                           DS.DA.EndTrans etRollback
                                           GoTo NextIteration
                                           'GoTo ErrHandler
                                       End If
                                   Else
                                       For j = 0 To VVP.DS.HowMany - 1
                                           varIDProdotto = VVP.DS.ListItem(0, j)
        
                                           DS.Fields("0").NewValue = DS.DA.GetCounter("anaganu0", "anuprogr", "anuidanu", "904", 0)
                                           DS.Fields("0").Updated = True
        
                                           DS.Fields("1").NewValue = varIDViaggio
                                           DS.Fields("1").Updated = True
        
                                           DS.Fields("2").NewValue = varIDConsegna
                                           DS.Fields("2").Updated = True
        
                                           DS.Fields("3").NewValue = varIDProdotto
                                           DS.Fields("3").Updated = True
        
                                           DS.Fields("4").NewValue = "SI"
                                           DS.Fields("4").Updated = True
        
                                           DS.Fields("5").NewValue = varDataEvento
                                           DS.Fields("5").Updated = True
        
                                           DS.Fields("9").NewValue = "K"
                                           DS.Fields("9").Updated = True
        
                                           objLogger.logInfo "Creazione dell'evento in ResOil."
                                           If Not Insert Then
                                               objLogger.logError "Insert ResOil evento."
                                               DS.DA.EndTrans etRollback
                                               GoTo NextIteration
                                               'GoTo ErrHandler
                                           End If
        
                                           ' Carica i prodotti della consegna e aggiorna i dati
                                           With VVP
                                               .DS.Fields.ResetFields
                                               .DS.Fields("0").FindValue = varIDViaggio
                                               .DS.Fields("1").FindValue = varIDConsegna
                                               .DS.Fields("2").FindValue = varIDProdotto
        
                                               If Not .Find Then
                                                   objLogger.logError "Find Consegna Prodotto."
                                                   DS.DA.EndTrans etRollback
                                                   GoTo NextIteration
                                                   'GoTo ErrHandler
                                               End If
                                               If .DS.HowMany = 0 Then
                                                   objLogger.logError "Prodotto Consegna non trovato"
                                                   DS.DA.EndTrans etRollback
                                                   GoTo NextIteration
                                                   'GoTo ErrHandler
                                               End If
        
                                               ' Aggiorna lo stato dello scarico (VVPFLSTA)
                                               .DS.Fields("37").NewValue = "2"
                                               .DS.Fields("37").Updated = True
                                               LogFieldValue .DS.Fields("37")
        
                                               ' Aggiorna il flag Modificato (VVPFLST5)
                                               .DS.Field("Stato Riga 5").NewValue = "S"
                                               .DS.Field("Stato Riga 5").Updated = True
                                               LogFieldValue .DS.Field("Stato Riga 5")
        
                                               '-- DP 05/06/08
                                               '-- Aggiorno il campo VVPFLAGG, che contiene l'indicazione dell'ultimo processo che ha aggiornato il record
                                               .DS.Fields("43").NewValue = "ME"
                                               .DS.Fields("43").Updated = True
                                               LogFieldValue .DS.Fields("43")
        
                                               ' Aggiorna l'ora inizio scarico (VVPHRINC)
                                               If blnUsaDatiGPS Then
                                                   .DS.Fields("38").NewValue = varOraEvento
                                                   .DS.Fields("38").Updated = True
                                                   LogFieldValue .DS.Fields("38")
                                               End If
        
                                               ' Salva
                                               If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
                                                   objLogger.logError "Save Consegna Prodotto."
                                                   DS.DA.EndTrans etRollback
                                                   GoTo NextIteration
                                                   'GoTo ErrHandler
                                               End If
                                               DS.DA.EndTrans etCommit
                                           End With
                                       Next j
                                   End If
        
                                   ' Aggiorna il viaggio
                                   With VVT
                                       If .CheckViaggioCompletato(varIDViaggio) Then
                                           If Not .SetStato("C") Then
                                               objLogger.logError "Set Viaggio Completato."
                                               DS.DA.EndTrans etRollback
                                               GoTo NextIteration
                                               'GoTo ErrHandler
                                           End If
                                       End If
        
                                       If blnUsaDatiGPS Then
                                           ' Aggiorna il flag trasmissione viaggio
                                           .DS.Fields("41").NewValue = "S"
                                           .DS.Fields("41").Updated = True
                                           LogFieldValue .DS.Fields("41")
        
                                           '-- DP 05/06/08
                                           '-- Aggiorno il campo VVTFLAGG, che contiene l'indicazione dell'ultimo processo che ha aggiornato il record
                                           .DS.Fields("52").NewValue = "ME"
                                           .DS.Fields("52").Updated = True
                                           LogFieldValue .DS.Fields("52")
                                       End If
        
                                       ' Salva
                                       If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
                                           objLogger.logError "Save Viaggio."
                                           DS.DA.EndTrans etRollback
                                           GoTo NextIteration
                                           'GoTo ErrHandler
                                       End If
        
                                       .DS.DA.EndTrans etCommit
                                   End With
'                               Else
'                                   objLogger.logError "Ora di reale di consegna già trasmessa su GES"
'                               End If
        
                               'DS.DA.EndTrans etCommit
        
                           Case Else
                               objLogger.logError "Tipo evento sconosciuto."
                               DS.DA.EndTrans etRollback
                               'GoTo NextIteration
                               'GoTo ErrHandler
        
                       End Select
                   Else
                       DS.DA.LogExecTime "WARN: Viaggio non trovato!", -1
                   End If
        
                   'DS.DA.EndTrans etCommit
                   

'comprobamos que el viaje sea de cepsa antes de generar el GES
'If Not (DS.DA.QueryCommand("SELECT VVTIDASC From VIAGVVT0 WHERE (VVTIDVVT = ?)", rowcep, lngcep, varIDViaggio)) Then
'End If
'Dim contcep As String
'
'If (lngcep > 0) Then
'   contcep = rowcep(0, 0)
'Else
'   If varIdevento = "001" Or varIdevento = "007" Then
'   contcep = "038"
'   Else
'   contcep = ""
'   End If
'End If


' Georeferenciar TODOS los clientes que no son Cepsa
If (contcep <> "038") Then


                      Select Case varIdevento
                      
                      Case "005"
                       
                        Dim Matclick, Mattr, MatCis, DatePos, sqlq As String
                        Dim DiffHrPos As Long
                        HREv = Format(HREv, "hh:mm")
                        Dim añoc As String
                        añoc = Mid(DYEv, 1, 4)
                        Dim mesc As String
                        mesc = Mid(DYEv, 5, 2)
                        Dim diac As String
                        diac = Mid(DYEv, 7, 2)
                        Dim FechaEvComp As Date
                        FechaEvComp = CDate(diac & "/" & mesc & "/" & añoc & " " & HREv)
                      
                        Dim VarKbufferPOSCLI As Variant
                        Dim lngPOSCLI As Long
                      
                            If DS.DA.QueryCommand(SQL_POS_CLI, VarKbufferPOSCLI, lngPOSCLI, clienterec) Then
                                 
                                      If ((Not IsEmpty(VarKbufferPOSCLI(0, 0)) And (Not IsEmpty(VarKbufferPOSCLI(1, 0)) And (VarKbufferPOSCLI(0, 0) <> 0)) And (VarKbufferPOSCLI(0, 0) <> 0) And (VarKbufferPOSCLI(0, 0) <> "") And (VarKbufferPOSCLI(0, 0) <> ""))) Then
                                            CordX = CStr(VarKbufferPOSCLI(0, 0))
                                            CordY = CStr(VarKbufferPOSCLI(1, 0))
                                     Else
                                        If (CordY = 0 Or CordX = 0) Then
                                                               
                                        ' Posicion tecnocom
                '                            If Not DS.DA.QueryCommand("SELECT ANAGAAM0.AAMIDDEU FROM VIAGVVT0 INNER JOIN  ANAGAAM0 ON VIAGVVT0.VVTIDAM1 = ANAGAAM0.AAMIDAAM WHERE     (VIAGVVT0.VVTIDVVT = ?)", Varbuff, 1, VarIDViag) Then
                '                            End If
                '                            If Not IsEmpty(Varbuff) Then idvehi = Varbuff(0, 0)
                '                            If Not DS.DA.QueryCommand("SELECT TOP 1 IDVEHIC, Latitud, Longitud, FEUTCVAL From PosicionesOesia WHERE (IDVEHIC = ?) AND (FEUTCVAL < ?) ORDER BY FEUTCVAL DESC", Varbuff, 1, idvehi, FechaEvComp) Then
                '                            End If
                                        '            Posicion Click & Find
                                        Dim Varbuff As Variant
                                        If Not DS.DA.QueryCommand("SELECT VVTIDAM1, VVTIDAM2 FROM VIAGVVT0 WHERE (VVTIDVVT = ?)", Varbuff, 1, varIDViaggio) Then
                                        End If
                                        If Not IsEmpty(Varbuff) Then
                                            Mattr = Varbuff(0, 0)
                                            If Not IsNull(Varbuff(1, 0)) Then
                                                If Not IsEmpty(Varbuff(1, 0)) And Varbuff(1, 0) <> "" Then MatCis = Varbuff(1, 0)
                                            End If
                                        End If
                                        
                                        Matclick = Mattr
                                        
                                        If Not IsNull(MatCis) Then
                                            If Not IsEmpty(MatCis) And MatCis <> "" Then Matclick = MatCis
                                        End If
                                        
                                        If Not DS.DA.QueryCommand("SELECT TOP (1) MATRICULA, LATITUD, LONGITUD, FECHAVAL FROM PosicionesClick WHERE (MATRICULA = ?) AND (FECHAVAL < ?) ORDER BY FECHAVAL DESC", Varbuff, 1, Matclick, FechaEvComp) Then
                                        End If '
                                        
                                            If Not IsEmpty(Varbuff) Then
                                                If Not IsNull(Varbuff(1, 0)) Then CordX = Varbuff(1, 0)
                                                If Not IsNull(Varbuff(2, 0)) Then CordY = Varbuff(2, 0)
                                                If Not IsNull(Varbuff(3, 0)) Then DatePos = Varbuff(3, 0)
                                            End If
                                            If DatePos = "" Then DatePos = "01/01/1900 0:00:00"
                                            DiffHrPos = DateDiff("n", DatePos, FechaEvComp)
                                         End If
                                        'Geolocalizamos el cliente
                                        sqlq = "UPDATE ANAGADZ0 SET ADZCGPSX = '" + CordX + "' ,ADZCGPSY = '" + CordY + "' ,ADZCOOCA = '" + CStr(DiffHrPos) + "' WHERE ADZIDADZ = '" + clienterec + "' "
                                        If Not DS.DA.ExecuteQuery(sqlq) Then
                                        End If
                                    End If
                            End If
                                      
                      
                      
                      End Select
                      
End If


If (contcep = "038") Then


                      Select Case varIdevento
                   
                        Case "001"
                          If (s = 0 And i = 0) Then
                                  If ListaGES(varIDViaggio, varIdevento, DYEv, HREv, CordX, CordY, , IDPalm) Then
                                  End If
                          End If
                                                
                        Case "002"
                            If (s = 0 And i = 0) Then
                                  If ListaGES(varIDViaggio, varIdevento, DYEv, HREv, CordX, CordY, , IDPalm) Then
                                  End If
                            End If
                                                   
                        Case "003"
                            If (s = 0 And i = 0) Then
                                  If ListaGES(varIDViaggio, varIdevento, DYEv, HREv, CordX, CordY, , IDPalm) Then
                                  End If
                            End If

                            ' Si tratta di un carico
                                varIDConsegna = vbNullString
                                varIDProdotto = vbNullString
                                strTipoEvento = "SAL"
                                
                        Case "005"
                          If (s = 0 And i = 0) Then
                                  If (esdoble) Then
                                  'Marcar el resoil el viaje como enviado en el caso de que sea un viaje con doble descarga.
                                        Dim Sql_update_dobles As String
                                        Sql_update_dobles = "UPDATE VIAGVVC0 SET VVCFLST5= 'S', viagvvc0.vvcflsta= 2 , viagvvc0.vvcflagg='ME' , VVCDTARR= '" + varDataEvento + "' , VVCHRARR= '1/1/1900 " + HREv + "' WHERE VVCIDVVT = " + varIDViaggio + " AND VVCIDADZ = '" + clienterec + "'"
                                        If Not DS.DA.ExecuteQuery(Sql_update_dobles) Then
                                        End If
                                     With VVT
                                     If .CheckViaggioCompletato(varIDViaggio) Then
                                         If Not .SetStato("C") Then
'                                             objLogger.logError "Set Viaggio Completato."
'                                             DS.DA.EndTrans etRollback
'                                             'GoTo NextIteration
'                                             'GoTo ErrHandler
                                         End If
                                     End If
                                     End With
                                  End If
                                  If ListaGES(varIDViaggio, varIdevento, DYEv, HREv, CordX, CordY, clienterec, IDPalm, , , eventooesia) Then
                                  End If
 
                          End If
                                ' Non è un carico, cerchiamo tra gli scarichi
                                If Not (DS.DA.QueryCommand(SQL_PALM_GET_CONSEGNA, varBufferConsegna, lnghowmanycons, eventooesia, varIDViaggio)) Then
                                        objLogger.logError "Query SQL_PALM_GET_CONSEGNA."
                                        GoTo ErrHandler
                                End If
                               varIDConsegna = varBufferConsegna(0, 0)
                               strTipoEvento = "INI"
                           
                        Case "006"
                             If (s = 0 And i = 0) Then
                                  If ListaGES(varIDViaggio, varIdevento, DYEv, HREv, CordX, CordY, clienterec, IDPalm) Then
                                  End If
                             End If
                        Case "007"
                              If (s = 0 And i = 0) Then
                                  If ListaGES(varIDViaggio, varIdevento, DYEv, HREv, CordX, CordY, , IDPalm, varDataEvento) Then
                                  End If
                              End If
                        Case "008"
                                If IsNull(varKBuffer(1, 0)) Then
                                    IdTarea = "0"
                                Else
                                    IdTarea = varKBuffer(1, 0)
                                End If
                                Orden = varKBuffer(5, 0)
                                If (s = 0 And i = 0) Then
                                      If ListaGES(varIDViaggio, varIdevento, DYEv, HREv, CordX, CordY, , IDPalm, , Orden, IdTarea) Then
                                      End If
                                End If
                        Case "999"
                            If (s = 0 And i = 0) Then
                                  If ListaGES(varIDViaggio, varIdevento, DYEv, HREv, CordX, CordY, clienterec) Then
                                  End If
                            End If
                            
                        Case Else
                       
                          'If (CInt(varIdevento) > 8 And CInt(varIdevento) < 21) Then
                          ' JG 25/11/21 Nuevos eventos incidencias
                          If ((CInt(varIdevento) > 8 And CInt(varIdevento) < 26)) Or (CInt(varIdevento) = 34) Then
                            
                            Orden = varKBuffer(5, 0)
                            If (s = 0 And i = 0) Then
                               If ListaGES(varIDViaggio, varIdevento, DYEv, HREv, CordX, CordY, clienterec, IDPalm, , Orden, , IdRegistro, ProdObserva) Then
                               End If
                            End If
                            
                          Else
                        
                           objLogger.logError "Errore durante la ricerca del Viaggio, Autista " & varIDAutista & ", Data " & varDataEvento
                           DS.DA.EndTrans etRollback
                           'GoTo NextIteration
                           
                          End If
        
                           
              
                   End Select
                   
End If
                   

               'Next i

               Next s
           End If

NextIteration:
        'Dim evoes As New CDataAccess
        '
        '    If Not (evoes.ExecuteQuery("UPDATE EventosMBFlotasmanual Set Capturado = 1 WHERE (IdRegistro = '" & Trim(idregistro) & "')")) Then
        '    End If
           Dim SQL_UPDATE_MBFlotas As String
           'Updateamos el registro para saber que esta recibido.
            SQL_UPDATE_MBFlotas = "UPDATE EventosMBFlotasmanual Set Capturado = 1 WHERE (IdRegistro = ?)"
           
           If Not DS.DA.ExecuteCommand(SQL_UPDATE_MBFlotas, IdRegistro) Then
           End If
                      
            If Not DS.DA.QueryCommand(SQL_GET_EVENTOS_PALM_TOTAL, varKBuffer, lngHowMany) Then
            objLogger.logError "Query SQL_GET_EVENTOS_PALM_TOTAL."
            GoTo ErrHandler
            End If

    Loop
    ScanEvents_Palm_nuevo = True

    Set VVP = Nothing
    Set VVT = Nothing
    Set VVC = Nothing

    objLogger.logDebug "Fine ScanEvents_palm senza errori."
    GoTo EndProc

ErrHandler:
    DS.DA.EndTrans etRollback

    Set VVP = Nothing
    Set VVT = Nothing
    Set VVC = Nothing

    objLogger.logError "Fine ScanEvents_SMS con ERRORI."

    ScanEvents_Palm_nuevo = False

EndProc:


End Function
Function CrearInsert(ByVal campo As String, ByVal values As String)
    
    Dim x As String
    
    If (Len(SQL_GES_campos) > 0) Then
        SQL_GES_campos = SQL_GES_campos + "," + campo
        SQL_GES_values = SQL_GES_values + "," + "'" + values + "'"
    Else
        SQL_GES_campos = campo
        SQL_GES_values = "'" + values + "'"
    End If
    
End Function

Function EliminaTareaPalm(IDORD As String, DataConnect As CConnect)
         
        
                       
                        Set Service = New VLocalize.VLocalize
                        Service.serviceLocation = "http://mb.oesia.com/mbflotastph/mbws/mbws.asmx"
                        Service.serviceTimeOut = 60
                        Service.serviceUsername = "wstph"
                        Service.servicePassword = "swtph"
                        Service.setServerID = DataConnect.ConnServer
                        Service.setServerPassword = DataConnect.ConnPWD
                        Service.setServerUser = DataConnect.ConnUserName
                         
                        If (Service.CancelaTarea(IDORD)) Then
                            MsgBox ("Orden eliminada")
                            Else
                            MsgBox ("Orden no encontrada")
                            End If
                       
                     If Not Service.isOpen Then
                           Service.Open
                       End If
                       If Service.isOpen Then
                           Service.Close
                       End If
End Function

Function EliminaServicioPalm(IDViag As String, DataConnect As CConnect)
                        
                        
                        Set Service = New VLocalize.VLocalize
                            
                       
                        Service.serviceLocation = "http://mb.oesia.com/mbflotastph/mbws/mbws.asmx"
                        Service.serviceTimeOut = 60
                        Service.serviceUsername = "wstph"
                        Service.servicePassword = "swtph"
                        Service.setServerID = DataConnect.ConnServer
                        Service.setServerPassword = DataConnect.ConnPWD
                        Service.setServerUser = DataConnect.ConnUserName
                        
                        If Not Service.isOpen Then
                                 Service.Open
                             End If
                        If (Service.DelServicioPalm(IDViag)) Then
                            MsgBox ("Servicio Eliminado")
                            Else
                            MsgBox ("Servicio no encontrado")
                            End If
                        If Service.isOpen Then
                            Service.Close
                        End If
End Function

Function ComprobarEstado(Refcep As String) As Boolean
Dim SQL_Estado As String
Dim Rowbuf As Variant
Dim lngbuf As Long

ComprobarEstado = False

Refcep = Left(Refcep, 10)

SQL_Estado = "SELECT     VIAGVVC0.VVCFLST5 FROM  VIAGVVC0 INNER JOIN VIAGVVP0 ON VIAGVVC0.VVCIDVVT = VIAGVVP0.VVPIDVVT AND VIAGVVC0.VVCIDVVC = VIAGVVP0.VVPIDVVC INNER JOIN ORDIOCT0 ON VIAGVVC0.VVCIDOCT = ORDIOCT0.OCTIDOCT WHERE  (ORDIOCT0.OCTRIFE1 = ?)"
If DS.DA.QueryCommand(SQL_Estado, Rowbuf, lngbuf, Refcep) Then
End If

If (lngbuf > 0) Then
    If Not (Rowbuf(0, 0) = "S") Then
        ComprobarEstado = True
    End If
End If
    
End Function
Function Genera004(VarIDViag As String, Cliente As String, DYEv As String, ByVal HREv As String, GPSX As String, GPSY As String)

Dim VarKbufferGes As Variant
Dim lngGes As Long
Dim i, j As Integer
Dim Varbuff As Variant
Dim columna, Value, sqlins As String
Dim FechaEvComp As Date
Dim DYEvN, HREvN As String

Dim SQL_GES_campos_sql(0 To 24) As String ' campos a rellenar del evento 004
SQL_GES_campos_sql(0) = "GESCODTR"
SQL_GES_campos_sql(1) = "GESNUMPD"
SQL_GES_campos_sql(2) = "GESNUMPR"
SQL_GES_campos_sql(3) = "GESCODDE"
SQL_GES_campos_sql(4) = "GESPRCOD"
SQL_GES_campos_sql(5) = "GESFACID"
SQL_GES_campos_sql(6) = "GESCANPR"
SQL_GES_campos_sql(7) = "GESUNMED"
SQL_GES_campos_sql(8) = "GESDYDES"
SQL_GES_campos_sql(9) = "GESCARTR"
SQL_GES_campos_sql(10) = "GESDESTR"
SQL_GES_campos_sql(11) = "GESDYTR0"
SQL_GES_campos_sql(12) = "GESHRTR0"
SQL_GES_campos_sql(13) = "GESVHIR0"
SQL_GES_campos_sql(14) = "GESVIAGR"
SQL_GES_campos_sql(15) = "GESCONR0"
SQL_GES_campos_sql(16) = "GESKMCEP"
SQL_GES_campos_sql(17) = "GESNUMVI"
SQL_GES_campos_sql(18) = "GESNOTES"
SQL_GES_campos_sql(19) = "GESKMCR0"
SQL_GES_campos_sql(20) = "GESDYACT"
SQL_GES_campos_sql(21) = "GESHRACT"
SQL_GES_campos_sql(22) = "GESHRR00"
SQL_GES_campos_sql(23) = "GESDYR00"


Dim añoc As String
añoc = Mid(DYEv, 1, 4)
Dim mesc As String
mesc = Mid(DYEv, 5, 2)
Dim diac As String
diac = Mid(DYEv, 7, 2)

FechaEvComp = CDate(diac & "/" & mesc & "/" & añoc & " " & HREv)

DYEvN = Format(FechaEvComp + TimeValue("00:05"), "yyyymmdd")
HREvN = Format(FechaEvComp + TimeValue("00:05"), "hh:mm")


If Not DS.DA.QueryCommand(SQL_GES_005006, VarKbufferGes, lngGes, VarIDViag, Cliente) Then
End If

If (lngGes = 0) Then
GoTo NextIteration
End If
     
For i = 0 To UBound(VarKbufferGes, 2)

'Añadimos la ultima posición que tenemos del vehiculo
'    If Not DS.DA.QueryCommand("SELECT ANAGAAM0.AAMIDDEU FROM VIAGVVT0 INNER JOIN  ANAGAAM0 ON VIAGVVT0.VVTIDAM1 = ANAGAAM0.AAMIDAAM WHERE     (VIAGVVT0.VVTIDVVT = ?)", Varbuff, 1, VarIDViag) Then
'    End If
'
'        If Not DS.DA.QueryCommand("SELECT TOP 1 IDVEHIC, Latitud, Longitud From PosicionesOesia WHERE (IDVEHIC = ?)", Varbuff, 1, CStr(Varbuff(0, 0))) Then
'        End If
'
        If Not CrearInsert("GESCORDX", GPSX) Then  ' CORDENADA
        End If
        If Not CrearInsert("GESCORDY", GPSY) Then ' CORDENADA
        End If
    
    
    If Not CrearInsert("GESCODEV", "004") Then ' CODIGOEVENTO
    End If
    If Not CrearInsert("GESDYEV0", DYEvN) Then ' Dia Evento
    End If
    HREv = Format(HREv, "hh:mm")
    If Not CrearInsert("GESHREV0", HREvN) Then ' Hora Evento
    End If
    If Not CrearInsert("IDVVT", VarIDViag) Then ' CORDENADA
    End If
    If Not CrearInsert("GESTRCEP", 0) Then
    End If

    
    
    'Ahora lo hace hasta el 23 por que el 24 es un campo que no va en VIAGGES0
     For j = 0 To 23  'UBound(VarKbufferGes, 1)
        If (IsNull(VarKbufferGes(j, i))) Then
         Value = ""
        Else
         Value = VarKbufferGes(j, i)
         'Multiplicamos la distancia por 2 para el evento 004 para reportar el retorno 15/4/15
         'If j = 19 Then Value = Value * 2
        
        'No ponemos distancia en evento 004
        
        If j = 19 Then Value = 0
         
        End If
        columna = SQL_GES_campos_sql(j)
        
        ' JG 12/6/2017 Solo ponemos distancias en los eventos 005 y 006

        If (SQL_GES_campos_sql(j) = "GESKMCR0") Then Value = "0"
                        
        If Not CrearInsert(columna, Value) Then
        End If
    Next j

    sqlins = "INSERT INTO VIAGGES0 (" + SQL_GES_campos + ") VALUES (" + SQL_GES_values + ")"
    DS.DA.ExecuteQuery (sqlins)
    SQL_GES_values = ""
    SQL_GES_campos = ""

Next i

NextIteration:

End Function

Function ListaGES(VarIDViag As String, CodEv As String, DYEv As String, HREv As String, CordX As String, CordY As String, Optional Cliente As String, Optional IDPalm As String, Optional ByVal Diaevento As String, Optional Orden As String, Optional IdTarea As String, Optional IdRegistro As String, Optional ObservaInc As String)


Dim VarKbufferGes As Variant
Dim VarKbufferGes001 As Variant
Dim lngGes As Long
Dim lngGes001 As Long
Dim VarKbufferCU As Variant
Dim lngGesCU As Long
Dim VarKbufferHC As Variant
Dim lngGesHC As Long
Dim buffpda As Variant
Dim lngpda As Long
Dim buffCond As Variant
Dim lngCond As Long
Dim buffSigVi As Variant
Dim lngSigVi As Long
Dim SQL_GEN_EV As String
Dim VarKbufferPOSFAC As Variant
Dim lngPOSFAC As Long
Dim VarKbufferPOSCLI As Variant
Dim lngPOSCLI As Long
Dim Varbuff As Variant
Dim idvehi As String
Dim buff01 As Variant
Dim lng01 As Long
Dim buff001 As Variant
Dim lng001 As Long
Dim buff02 As Variant
Dim lng02 As Long
Dim buff03 As Variant
Dim lng03 As Long
Dim buff05 As Variant
Dim lng05 As Long
Dim rowev As Variant
Dim lngev As Long
Dim var004 As Variant
Dim lng004 As Long
Dim FechaEvComp As Date
Dim DYEvN, HREvN As String
Dim j, i, p As Integer
Dim columna, Value As String
Dim sqlins As String
Dim sqlins004 As String
Dim CUCarga As String
Dim sqlq As String
Dim SQL_GES_campos_sql(0 To 24) As String ' campos del evento 002,003,004
Dim TmpAl As Long
Dim FechaSalida, FechaLlegada As Date
Dim Motivo As String
Dim PararEvento As Boolean
Dim Gen_Por As String
Dim Parar_Datos As Boolean
Dim DiffHrPos As Long
Dim DatePos As String
Dim Mattr As String
Dim MatCis As String
Dim Matclick As String
Dim Posfactx As String
Dim Posfacty As String
Dim SQL_INSERTA_GIS As String
Dim FechaPosVal As String
Dim nuevafepos As String
Dim ValorKm As Integer


SQL_GES_campos_sql(0) = "GESCODTR"
SQL_GES_campos_sql(1) = "GESNUMPD"
SQL_GES_campos_sql(2) = "GESNUMPR"
SQL_GES_campos_sql(3) = "GESCODDE"
SQL_GES_campos_sql(4) = "GESPRCOD"
SQL_GES_campos_sql(5) = "GESFACID"
SQL_GES_campos_sql(6) = "GESCANPR"
SQL_GES_campos_sql(7) = "GESUNMED"
SQL_GES_campos_sql(8) = "GESDYDES"
SQL_GES_campos_sql(9) = "GESCARTR"
SQL_GES_campos_sql(10) = "GESDESTR"
SQL_GES_campos_sql(11) = "GESDYTR0"
SQL_GES_campos_sql(12) = "GESHRTR0"
SQL_GES_campos_sql(13) = "GESVHIR0"
SQL_GES_campos_sql(14) = "GESVIAGR"
SQL_GES_campos_sql(15) = "GESCONR0"
SQL_GES_campos_sql(16) = "GESKMCEP"
SQL_GES_campos_sql(17) = "GESNUMVI"
SQL_GES_campos_sql(18) = "GESNOTES"
SQL_GES_campos_sql(19) = "GESKMCR0"
SQL_GES_campos_sql(20) = "GESDYACT"
SQL_GES_campos_sql(21) = "GESHRACT"
SQL_GES_campos_sql(22) = "GESHRR00"
SQL_GES_campos_sql(23) = "GESDYR00"
SQL_GES_campos_sql(24) = "VVTSLUDG"


Dim SQL_GES_campos_sql005006(0 To 24) As String ' campos del eveto 005,006
SQL_GES_campos_sql005006(0) = "GESCODTR"
SQL_GES_campos_sql005006(1) = "GESNUMPD"
SQL_GES_campos_sql005006(2) = "GESNUMPR"
SQL_GES_campos_sql005006(3) = "GESCODDE"
SQL_GES_campos_sql005006(4) = "GESPRCOD"
SQL_GES_campos_sql005006(5) = "GESFACID"
SQL_GES_campos_sql005006(6) = "GESCANPR"
SQL_GES_campos_sql005006(7) = "GESUNMED"
SQL_GES_campos_sql005006(8) = "GESDYDES"
SQL_GES_campos_sql005006(9) = "GESCARTR"
SQL_GES_campos_sql005006(10) = "GESDESTR"
SQL_GES_campos_sql005006(11) = "GESDYTR0"
SQL_GES_campos_sql005006(12) = "GESHRTR0"
SQL_GES_campos_sql005006(13) = "GESVHIR0"
SQL_GES_campos_sql005006(14) = "GESVIAGR"
SQL_GES_campos_sql005006(15) = "GESCONR0"
SQL_GES_campos_sql005006(16) = "GESKMCEP"
SQL_GES_campos_sql005006(17) = "GESNUMVI"
SQL_GES_campos_sql005006(18) = "GESNOTES"
SQL_GES_campos_sql005006(19) = "GESKMCR0"
SQL_GES_campos_sql005006(20) = "GESDYACT"
SQL_GES_campos_sql005006(21) = "GESHRACT"
SQL_GES_campos_sql005006(22) = "GESHRR00"
SQL_GES_campos_sql005006(23) = "GESDYR00"
SQL_GES_campos_sql005006(24) = "VVTSLUDG"




Dim SQL_GES_campos_sql_001007(0 To 4) As String
SQL_GES_campos_sql_001007(0) = "GESCODTR"
SQL_GES_campos_sql_001007(1) = "GESHRCAR"
SQL_GES_campos_sql_001007(2) = "GESCONR0"
SQL_GES_campos_sql_001007(3) = "GESVHIR0"
SQL_GES_campos_sql_001007(4) = "GESFACID"

'Iniciamos las variables
Parar_Datos = False
PararEvento = False
SQL_GES_values = ""
SQL_GES_campos = ""
HREv = Format(HREv, "hh:mm")
Dim añoc As String
añoc = Mid(DYEv, 1, 4)
Dim mesc As String
mesc = Mid(DYEv, 5, 2)
Dim diac As String
diac = Mid(DYEv, 7, 2)

FechaEvComp = CDate(diac & "/" & mesc & "/" & añoc & " " & HREv)

' Asignamos al campo Gen_por quien ha generado el evento M-Manual R-Resoil S-SMS P-PDA


Select Case IDPalm

    Case "0"
    Gen_Por = "M"
    
    Case "999"
    Gen_Por = "R"

    Case "666"
    Gen_Por = "S"
    
    Case Else
    Gen_Por = "P"
    
End Select


If (CodEv = "001") Then
                   
    If VarIDViag = "-1" Then
                   
        If DS.DA.QueryCommand(SQL_GES_001, VarKbufferGes001, lngGes001, IDPalm) Then
        If lngGes001 > 0 Then
        VarIDViag = CStr(VarKbufferGes001(0, 0))
                           
        If Not DS.DA.QueryCommand(SQL_GES_002003, VarKbufferGes, lngGes, VarIDViag) Then
        End If
         
        If (lngGes > 0) Then ' EL VIAJE PUEDE ESTAR ANULADO (EN ESE CASO OMITE Y SIGUE)
        
             'Calculo de Coeficiente de Utilización
              If Not DS.DA.QueryCommand(SQL_GES_CUCARGA, VarKbufferCU, lngGesCU, VarIDViag) Then
             End If
             
             If (lngGesCU = 0) Then
             CUCarga = "000.000"
             Else
                    If (IsNull(VarKbufferCU(0, 0)) Or IsNull(VarKbufferCU(1, 0)) Or VarKbufferCU(0, 0) = 0 Or VarKbufferCU(1, 0) < 1) Then
                        CUCarga = "090.55"
                        Else
                        CUCarga = CStr((VarKbufferCU(0, 0) / VarKbufferCU(1, 0)) * 100)
                        
                         CUCarga = Format$(CUCarga, "###.##")
                         If (CInt(CUCarga) < 0) Then
                            CUCarga = CStr(CInt(CUCarga) * -1)
                         End If
                         If (CInt(CUCarga) < 100 And CInt(CUCarga) > 10) Then
                            CUCarga = "0" + CUCarga
                         ElseIf (CInt(CUCarga) < 10) Then
                            CUCarga = "00" + CUCarga
                         End If
                     End If
             End If
            
                                         
             For i = 0 To UBound(VarKbufferGes, 2)
             
                 If DS.DA.QueryCommand(SQL_POS_FAC, VarKbufferPOSFAC, lngPOSFAC, CStr(VarKbufferGes(5, i))) Then
                 
                  If Not IsNull(VarKbufferPOSFAC(1, 0)) Then
                  Posfactx = (VarKbufferPOSFAC(1, 0))
                  Else
                  Posfactx = ""
                  End If
                  
                  If Not IsNull(VarKbufferPOSFAC(2, 0)) Then
                  Posfacty = (VarKbufferPOSFAC(2, 0))
                  Else
                  Posfacty = ""
                  End If
                 
                     If Not CrearInsert("GESCORDX", Posfactx) Then ' CORDENADA
                     End If
                     If Not CrearInsert("GESCORDY", Posfacty) Then ' CORDENADA
                     End If
                 Else
                     If Not CrearInsert("GESCORDX", CordY) Then ' CORDENADA
                     End If
                     If Not CrearInsert("GESCORDY", CordX) Then ' CORDENADA
                     End If
                 End If
                 If Not CrearInsert("GESCARR0", CUCarga) Then 'Coeficiente de Utilización
                 End If
                 If Not CrearInsert("GESDYEV0", DYEv) Then ' Dia Evento
                 End If
                 HREv = Format(HREv, "hh:mm")
                 If Not CrearInsert("GESHREV0", HREv) Then ' Hora Evento
                 End If
                 If Not CrearInsert("IDVVT", VarIDViag) Then ' CORDENADA
                 End If
                 'Ahora lo hace hasta el 23 por que el 24 es un campo que no va en VIAGGES0
                 For j = 0 To 23  'UBound(VarKbufferGes, 1)
                 
                        If (IsNull(VarKbufferGes(j, i))) Then
                         Value = ""
                        Else
                         Value = VarKbufferGes(j, i)
                        End If
                        'NO INSERTAMOS EL HORARIO SALVA 21/10/11 GESDYTR0
                        'If SQL_GES_campos_sql(j) = "GESDYTR0" Or SQL_GES_campos_sql(j) = "GESHRTR0" Then Value = ""
                        columna = SQL_GES_campos_sql(j)
                        
                        'JG 12/6/2017 Solo ponemos distancias en los eventos 002 y 005
                        If ((CodEv = "002") Or (CodEv = "006")) And (SQL_GES_campos_sql(j) = "GESKMCR0") Then
                        If Not IsNumeric(Value) Then Value = "0"
                        ValorKm = CInt(Value)
                        Value = CStr(Round(ValorKm * 1, 0))
                        Else
                        If (SQL_GES_campos_sql(j) = "GESKMCR0") Then Value = "0"
                        End If
                        
                        If Not CrearInsert(columna, Value) Then
                        End If
                 Next j
                     If Not CrearInsert("GESCODEV", CodEv) Then ' CODIGOEVENTO
                     End If
                ' Paramos el evento para rellenar compartimentos
                     If Not CrearInsert("GESTRCEP", 0) Then
                     End If
    
    'INSERT INTO VIAGES (SQL_CAMPOS) VALUES (SQL_VALUES)

                    sqlins = "INSERT INTO VIAGGES0 (" + SQL_GES_campos + ") VALUES (" + SQL_GES_values + ")"
                    DS.DA.ExecuteQuery (sqlins)
                    SQL_GES_values = ""
                    SQL_GES_campos = ""
                Next i
         End If
          
           End If
        End If
    
    Else
    ' Insertamos el 001 con el viaje que nos viene dado en el evento
        If Not DS.DA.QueryCommand(SQL_GES_002003, VarKbufferGes, lngGes, VarIDViag) Then
        End If
         
        If (lngGes > 0) Then ' EL VIAJE PUEDE ESTAR ANULADO (EN ESE CASO OMITE Y SIGUE)
        
             'Calculo de Coeficiente de Utilización
              If Not DS.DA.QueryCommand(SQL_GES_CUCARGA, VarKbufferCU, lngGesCU, VarIDViag) Then
             End If
             
             If (lngGesCU = 0) Then
             CUCarga = "000.000"
             Else
                    If (IsNull(VarKbufferCU(0, 0)) Or IsNull(VarKbufferCU(1, 0)) Or VarKbufferCU(0, 0) = 0 Or VarKbufferCU(1, 0) < 1) Then
                        CUCarga = "090.55"
                        Else
                        CUCarga = CStr((VarKbufferCU(0, 0) / VarKbufferCU(1, 0)) * 100)
                        
                         CUCarga = Format$(CUCarga, "###.##")
                         If (CInt(CUCarga) < 0) Then
                            CUCarga = CStr(CInt(CUCarga) * -1)
                         End If
                         If (CInt(CUCarga) < 100 And CInt(CUCarga) > 10) Then
                            CUCarga = "0" + CUCarga
                         ElseIf (CInt(CUCarga) < 10) Then
                            CUCarga = "00" + CUCarga
                         End If
                     End If
             End If
            
                                         
             For i = 0 To UBound(VarKbufferGes, 2)
             
                 If DS.DA.QueryCommand(SQL_POS_FAC, VarKbufferPOSFAC, lngPOSFAC, CStr(VarKbufferGes(5, i))) Then
                 
                  If Not IsNull(VarKbufferPOSFAC(1, 0)) Then
                  Posfactx = (VarKbufferPOSFAC(1, 0))
                  Else
                  Posfactx = ""
                  End If
                  
                  If Not IsNull(VarKbufferPOSFAC(2, 0)) Then
                  Posfacty = (VarKbufferPOSFAC(2, 0))
                  Else
                  Posfacty = ""
                  End If
                 
                     If Not CrearInsert("GESCORDX", Posfactx) Then ' CORDENADA
                     End If
                     If Not CrearInsert("GESCORDY", Posfacty) Then ' CORDENADA
                     End If
                 Else
                     If Not CrearInsert("GESCORDX", CordY) Then ' CORDENADA
                     End If
                     If Not CrearInsert("GESCORDY", CordX) Then ' CORDENADA
                     End If
                 End If
                 If Not CrearInsert("GESCARR0", CUCarga) Then 'Coeficiente de Utilización
                 End If
                 If Not CrearInsert("GESDYEV0", DYEv) Then ' Dia Evento
                 End If
                 HREv = Format(HREv, "hh:mm")
                 If Not CrearInsert("GESHREV0", HREv) Then ' Hora Evento
                 End If
                 If Not CrearInsert("IDVVT", VarIDViag) Then ' CORDENADA
                 End If
                 'Ahora lo hace hasta el 23 por que el 24 es un campo que no va en VIAGGES0
                 For j = 0 To 23  'UBound(VarKbufferGes, 1)
                        If (IsNull(VarKbufferGes(j, i))) Then
                         Value = ""
                        Else
                         Value = VarKbufferGes(j, i)
                        End If
                        columna = SQL_GES_campos_sql(j)
                        
                        'JG 12/6/2017 Solo ponemos distancias en los eventos 002 y 005
                        If ((CodEv = "002") Or (CodEv = "006")) And (SQL_GES_campos_sql(j) = "GESKMCR0") Then
                        If Not IsNumeric(Value) Then Value = "0"
                        ValorKm = CInt(Value)
                        Value = CStr(Round(ValorKm * 1, 0))
                        Else
                        If (SQL_GES_campos_sql(j) = "GESKMCR0") Then Value = "0"
                        End If
                        
                        'NO INSERTAMOS EL HORARIO SALVA 21/10/11 GESDYTR0
                        'If SQL_GES_campos_sql(j) = "GESDYTR0" Or SQL_GES_campos_sql(j) = "GESHRTR0" Then Value = ""

                        If Not CrearInsert(columna, Value) Then
                        End If
                 Next j
                     If Not CrearInsert("GESCODEV", CodEv) Then ' CODIGOEVENTO
                     End If
                ' Paramos el evento para rellenar compartimentos
                     If Not CrearInsert("GESTRCEP", 0) Then
                     End If
    
    
    'INSERT INTO VIAGES (SQL_CAMPOS) VALUES (SQL_VALUES)
                    sqlins = "INSERT INTO VIAGGES0 (" + SQL_GES_campos + ") VALUES (" + SQL_GES_values + ")"
                    DS.DA.ExecuteQuery (sqlins)
                    SQL_GES_values = ""
                    SQL_GES_campos = ""
                Next i
         End If
        
    End If
ElseIf (CodEv = "002" Or CodEv = "003") Then
            

              
     If Not DS.DA.QueryCommand(SQL_GES_002003, VarKbufferGes, lngGes, VarIDViag) Then
     End If
    
   If (lngGes > 0) Then ' EL VIAJE PUEDE ESTAR ANULADO (EN ESE CASO OMITE Y SIGUE)
     
   ' Controlamos que el viaje no se haya de parar y miramos que sea recibido por PDA o por SMS  en cuyo caso pasamos pararevento a true
   ' más adelante evaluaremos el control
     
   If VarKbufferGes(24, 0) = "N" Then Parar_Datos = True
     
    If Parar_Datos = True And (Gen_Por = "P" Or Gen_Por = "S") Then
    
    Motivo = "Evento parado manualmente"
    PararEvento = True
    
    End If
     
    Dim sql_ex05 As String
    sql_ex05 = "SELECT * FROM VIAGGES0 WHERE IDVVT = ? AND GESCODEV = '005'"
    If DS.DA.QueryCommand(sql_ex05, buff05, lng05, VarIDViag) Then
    End If
    
    If lng05 > 0 Then
    
        Dim añol As String
        añol = Mid(buff05(36, 0), 1, 4)
        Dim mesl As String
        mesl = Mid(buff05(36, 0), 5, 2)
        Dim dial As String
        dial = Mid(buff05(36, 0), 7, 2)
        FechaLlegada = CDate(dial & "/" & mesl & "/" & añol & " " & buff05(37, 0))
        If (Round((FechaEvComp - FechaLlegada) * 60 * 24) > 0) Then
            PararEvento = True
            Motivo = "Hay un evento 005 anterior."
        End If
    End If
     
     If (CodEv = "002") And PararEvento = False Then
        
        
        
        'Comprobamos que exista un evento 001 para este viaje
         
            Dim sql_ex01 As String
            sql_ex01 = "SELECT * FROM VIAGGES0 WHERE IDVVT = ? AND GESCODEV = '001'"
            If DS.DA.QueryCommand(sql_ex01, buff01, lng01, VarIDViag) Then
            End If

            If Not (lng01 > 0) Then 'SI YA HAY UN EVENTO 001 PARA ESTE VIAJE TODO CORRECTO, EN EL CASO CONTRARIO HAY QUE GENERAR UN EVENTO 001 PARA ESTE VIAJE.
                                ' le restamos 25 minutos al evento
                                TmpAl = Aleatorio(20, 30)
                                DYEvN = Format(FechaEvComp - TimeValue("00:" & TmpAl), "yyyymmdd")
                                HREvN = Format(FechaEvComp - TimeValue("00:" & TmpAl), "hh:mm:ss")
                                SQL_GEN_EV = "INSERT INTO EventosMBFlotasManual (IdUsuarioMB,CodigoEvento,FechaEvento,FechaServidor,EntaEnLoc,CodigoServicio, CodigoTarea) VALUES ('999','001','" + DYEvN + " " + HREvN + "',GETDATE (),'0','" & CStr(VarIDViag) & "','-1')"
                                If DS.DA.ExecuteQuery(SQL_GEN_EV) Then
                                End If
            End If
    End If
    
    
         'Calculo de Coeficiente de Utilización
          If Not DS.DA.QueryCommand(SQL_GES_CUCARGA, VarKbufferCU, lngGesCU, VarIDViag) Then
         End If
         
         If (lngGesCU = 0) Then
         CUCarga = "000.000"
         Else
                If (IsNull(VarKbufferCU(0, 0)) Or IsNull(VarKbufferCU(1, 0)) Or VarKbufferCU(0, 0) = 0 Or VarKbufferCU(1, 0) < 1) Then
                    CUCarga = "090.55"
                    Else
                    CUCarga = CStr((VarKbufferCU(0, 0) / VarKbufferCU(1, 0)) * 100)
                    
                     CUCarga = Format$(CUCarga, "###.##")
                     If (CInt(CUCarga) < 0) Then
                        CUCarga = CStr(CInt(CUCarga) * -1)
                     End If
                     If (CInt(CUCarga) < 100 And CInt(CUCarga) > 10) Then
                        CUCarga = "0" + CUCarga
                     ElseIf (CInt(CUCarga) < 10) Then
                        CUCarga = "00" + CUCarga
                     End If
                 End If
         End If
        
                                     
         For i = 0 To UBound(VarKbufferGes, 2)
         'Introducimos las cordenadas del evento, en el caso que tengamos georeferenciada la factoria  introducimos los datos que tenemos, en el caso
         'contrario introducimos las coordenadas del evento.

             If DS.DA.QueryCommand(SQL_POS_FAC, VarKbufferPOSFAC, lngPOSFAC, CStr(VarKbufferGes(5, i))) Then
               
                 
                  If Not IsNull(VarKbufferPOSFAC(1, 0)) Then
                    Posfactx = (VarKbufferPOSFAC(1, 0))
                  Else
                  Posfactx = ""
                  End If
                  
                  If Not IsNull(VarKbufferPOSFAC(2, 0)) Then
                  Posfacty = (VarKbufferPOSFAC(2, 0))
                  Else
                  Posfacty = ""
                  End If
                 
                 If Not CrearInsert("GESCORDX", Posfactx) Then ' CORDENADA
                 End If
                 If Not CrearInsert("GESCORDY", Posfacty) Then ' CORDENADA
                 End If
                 CordX = CStr(Posfactx)
                 CordY = CStr(Posfacty)
                 
                 ' Insertamos posiciones de la factoría para el vehículo para que no salten alertas GIS
                 If CodEv = "003" Then
                        ' Buscamos la matrícula del vehículo
                        If Not DS.DA.QueryCommand("SELECT VVTIDAM1, VVTIDAM2 FROM VIAGVVT0 WHERE (VVTIDVVT = ?)", Varbuff, 1, VarIDViag) Then
                        End If
                        If Not IsEmpty(Varbuff) Then
                            Mattr = Varbuff(0, 0)
                            If Not IsNull(Varbuff(1, 0)) Then
                                If Not IsEmpty(Varbuff(1, 0)) And Varbuff(1, 0) <> "" Then MatCis = Varbuff(1, 0)
                            End If
                        End If
                        
                        ' Mandar unas posiciónes en la factoría
'
'                        Matclick = Mattr
'                        If Not IsNull(MatCis) Then
'                            If Not IsEmpty(MatCis) And MatCis <> "" Then Matclick = MatCis
'                        End If
'
'                        nuevafepos = FechaEvComp
'                        FechaPosVal = FechaEvComp
'                        FechaPosVal = DateAdd("d", -1, FechaPosVal)
'                        For p = 0 To 2 ' Numero de posiciones a insertar
'                        nuevafepos = DateAdd("s", 30, nuevafepos)
'                        SQL_INSERTA_GIS = "Insert Into Posicionesclick (MATRICULA, FECHAVAL, FECHALLE, FECHAPOS, Latitud, Longitud) Values ('" & Matclick & "','" & FechaPosVal & "',getdate(),'" & nuevafepos & "','" & Posfactx & "','" & Posfacty & "')"
'                        If Not DS.DA.ExecuteQuery(SQL_INSERTA_GIS) Then
'                        End If
'                        Next p
                End If
             
             Else
                 If Not CrearInsert("GESCORDX", CordX) Then ' CORDENADA
                 End If
                 If Not CrearInsert("GESCORDY", CordY) Then ' CORDENADA
                 End If
             End If
             If Not CrearInsert("GESCARR0", CUCarga) Then 'Coeficiente de Utilización
             End If
             If Not CrearInsert("GESDYEV0", DYEv) Then ' Dia Evento
             End If
               HREv = Format(HREv, "hh:mm")
             If Not CrearInsert("GESHREV0", HREv) Then ' Hora Evento
             End If
             If Not CrearInsert("IDVVT", VarIDViag) Then ' CORDENADA
             End If
             'Bucle de generación de el insert
             'Ahora lo hace hasta el 23 por que el 24 es un campo que no va en VIAGGES0
                 For j = 0 To 23  'UBound(VarKbufferGes, 1)
                 If (IsNull(VarKbufferGes(j, i))) Then
                  Value = ""
                 Else
                  Value = VarKbufferGes(j, i)
                 End If
                 columna = SQL_GES_campos_sql(j)
                 

                 'NO INSERTAMOS EL HORARIO SALVA 21/10/11 GESDYTR0
                'If (CodEv = "002") And (SQL_GES_campos_sql(j) = "GESDYTR0" Or SQL_GES_campos_sql(j) = "GESHRTR0") Then Value = ""
                
                ' No ponemos distancias en los eventos 003-004-006
                'If ((CodEv = "003") Or (CodEv = "004") Or (CodEv = "006")) And (SQL_GES_campos_sql(j) = "GESKMCR0") Then Value = "0"

                'JG 12/6/2017 Solo ponemos distancias en los eventos 002 y 005
                If ((CodEv = "002") Or (CodEv = "006")) And (SQL_GES_campos_sql(j) = "GESKMCR0") Then
                If Not IsNumeric(Value) Then Value = "0"
                ValorKm = CInt(Value)
                Value = CStr(Round(ValorKm * 1, 0))
                Else
                If (SQL_GES_campos_sql(j) = "GESKMCR0") Then Value = "0"
                End If
                
                
                
                 If Not CrearInsert(columna, Value) Then
                 End If
             Next j
             
                 If Not CrearInsert("GESCODEV", CodEv) Then ' CODIGOEVENTO
                 End If
                 
                ' Paramos el evento para rellenar compartimentos
                 If Not CrearInsert("GESTRCEP", 0) Then
                 End If


                
                'INSERT INTO VIAGES (SQL_CAMPOS) VALUES (SQL_VALUES)
                If PararEvento = True Then
                    
                    If Not CrearInsert("GESPARAR", 1) Then
                    End If
                    If Not CrearInsert("GESOBCEP", Motivo) Then
                    End If
                 
                 End If
                    
                    sqlins = "INSERT INTO VIAGGES0 (" + SQL_GES_campos + ") VALUES (" + SQL_GES_values + ")"
                    DS.DA.ExecuteQuery (sqlins)
                    SQL_GES_values = ""
                    SQL_GES_campos = ""
                    
                    
                
             If (CodEv = "003" And i = 0) And PararEvento = False Then
                'buscamos el primer cliente del viaje
                If Not DS.DA.QueryCommand("SELECT TOP 1 VVCIDADZ From VIAGVVC0 WHERE (VVCIDVVT = ?) ORDER BY VVCIDVVC", var004, lng004, VarIDViag) Then
                End If
                'Generamos  004
                If Genera004(VarIDViag, CStr(var004(0, 0)), DYEv, HREv, CordX, CordY) Then
                End If
                
                'Comprobamos que exista un evento 002 para este viaje
                 
                    Dim sql_ex02 As String
                    sql_ex02 = "SELECT * FROM VIAGGES0 WHERE IDVVT = ? AND GESCODEV = '002'"
                    If DS.DA.QueryCommand(sql_ex02, buff02, lng02, VarIDViag) Then
                    End If
        
                    If Not (lng02 > 0) Then 'SI YA HAY UN EVENTO 002 PARA ESTE VIAJE TODO CORRECTO, EN EL CASO CONTRARIO HAY QUE GENERAR UN EVENTO 002 PARA ESTE VIAJE.
                                        ' le restamos 25 minutos al evento
                                        TmpAl = Aleatorio(20, 35)
                                        DYEvN = Format(FechaEvComp - TimeValue("00:" & TmpAl), "yyyymmdd")
                                        HREvN = Format(FechaEvComp - TimeValue("00:" & TmpAl), "hh:mm:ss")
                                        SQL_GEN_EV = "INSERT INTO EventosMBFlotasManual (IdUsuarioMB,CodigoEvento,FechaEvento,FechaServidor,EntaEnLoc,CodigoServicio, CodigoTarea) VALUES ('999','002','" + DYEvN + " " + HREvN + "',GETDATE (),'0','" & CStr(VarIDViag) & "','-1')"
                                        If DS.DA.ExecuteQuery(SQL_GEN_EV) Then
                                        End If
                           
                    End If
                
            End If
            
            Next i
     End If

ElseIf (CodEv = "005" Or CodEv = "006") Then



  If Not DS.DA.QueryCommand(SQL_GES_005006, VarKbufferGes, lngGes, VarIDViag, Cliente) Then
  End If

   ' Controlamos que el viaje no se haya de parar y miramos que sea recibido por PDA o por SMS  en cuyo caso pasamos pararevento a true
   ' más adelante evaluaremos el control


   If VarKbufferGes(24, 0) = "N" Then Parar_Datos = True
     
   If Parar_Datos = True And (Gen_Por = "P" Or Gen_Por = "S") Then
    
    Motivo = "Evento parado manualmente"
    PararEvento = False
    
   End If

        If (CodEv = "006") Then
        'DESPUES DE UN EVENTO 006 EN EL CASO DE VOLVER A CARGAR HAY QUE INFORMAR A CEPSA
               Dim Sql_next As String
               Sql_next = "SELECT * FROM VIAGVVC0 WHERE  (VVCFLST5 <> 'S' OR  VVCFLST5 IS NULL) AND (VVCIDVVT= " + VarIDViag + ")"
               If Not DS.DA.QueryCommand(Sql_next, buff01, lng01) Then 'Buscamos si ha acabado el viaje
               End If
               If (lng01 = 0) Then
               'en el caso de haya acabado el viaje buscamos que conductor ha hecho el viaje
                     If Not DS.DA.QueryCommand("SELECT VVTIDAU1 FROM VIAGVVT0 WHERE (VVTIDVVT = ?)", buffCond, lngCond, VarIDViag) Then
                     End If
                       If (lngCond > 0 And Not IsNull(buffCond(0, 0))) Then
                        If Not DS.DA.QueryCommand("SELECT TOP 1 VVTIDVVT FROM VIAGVVT0 WHERE (VVTIDAU1 = ?) AND (VVTFLST3 <> 'S' OR VVTFLST3 IS NULL) AND (VVTDTINC > GETDATE() - 1) AND (VVTANNUL = 0) AND (VVTDTINC < GETDATE()) ORDER BY VVTHRINC", buffSigVi, lngSigVi, CStr(buffCond(0, 0))) Then
                        End If
                        
                        ' JG 1/12/20 Comprobamos que el siguiente viaje no tenga ya un evento 001
                        If lngSigVi > 0 Then
                            Dim SLQ_001env As String
                            SLQ_001env = "SELECT * FROM VIAGGES0 WHERE IDVVT = ? AND GESCODEV = '001'"
                            If DS.DA.QueryCommand(SLQ_001env, buff001, lng001, CStr(buffSigVi(0, 0))) Then
                            End If
                        End If
                        
                        If lngSigVi > 0 And Not lng001 > 0 Then 'SI YA HAY UN EVENTO 001 PARA ESTE VIAJE NO SE ENVIA EL EVENTO 001.
                        'If lngSigVi > 0 Then
                                ' si tiene más viajes Insertamos un evento 001 para el siguiente viaje
                                DYEvN = Format(FechaEvComp + TimeValue("00:01"), "yyyymmdd")
                                HREvN = Format(FechaEvComp + TimeValue("00:01"), "hh:mm:ss")
                                SQL_GEN_EV = "INSERT INTO EventosMBFlotasManual (IdUsuarioMB,CodigoEvento,FechaEvento,FechaServidor,EntaEnLoc,CodigoServicio, CodigoTarea) VALUES ('999','001','" + DYEv + " " + HREv + "',GETDATE (),'0','" + CStr(buffSigVi(0, 0)) + "','-1')"
                               If DS.DA.ExecuteQuery(SQL_GEN_EV) Then
                               End If
                        End If
                      End If
               End If
        End If
                         
                'Calculo de Coeficiente de Utilización
                 If Not DS.DA.QueryCommand(SQL_GES_CUCARGA, VarKbufferCU, lngGesCU, VarIDViag) Then
                End If
               If (lngGesCU = 0) Then
                CUCarga = "000.000"
                Else
                    If (IsNull(VarKbufferCU(0, 0)) Or IsNull(VarKbufferCU(1, 0)) Or VarKbufferCU(0, 0) = 0 Or VarKbufferCU(1, 0) < 1) Then
                         CUCarga = "090.55"
                     Else
                         CUCarga = CStr((VarKbufferCU(0, 0) / VarKbufferCU(1, 0)) * 100)
                     
                        CUCarga = Format$(CUCarga, "###.##")
                        If (CInt(CUCarga) < 100 And CInt(CUCarga) > 10) Then
                           CUCarga = "0" + CUCarga
                        ElseIf (CInt(CUCarga) < 10) Then
                           CUCarga = "00" + CUCarga
                        End If
                     End If
                End If
                
        If (lngGes = 0) Then
            GoTo NextIteration
        End If
        
            For i = 0 To UBound(VarKbufferGes, 2)
           
                 'Buscar Horario cliente para poder parar la transmision a cepsa si se desea.
                'Comprobamos que la opcion de parar este activa
               If (CodEv = "005") Then
               
                Dim desde1 As Date
                Dim desde2 As Date
                Dim desde3 As Date
                Dim hasta1 As Date
                Dim hasta2 As Date
                Dim hasta3 As Date
                Dim margen As Date
                Dim horafinal As Date

                ' Establecemos en margen el margen de horarario que nos dan para cumplir
                margen = "01:00:00"
                
                      If DS.DA.QueryCommand(SQL_GES_HorariosCli, VarKbufferHC, lngGesHC, VarIDViag, Cliente) Then
                      If Not IsNull(VarKbufferHC(0, 0)) Then desde1 = VarKbufferHC(0, 0)
                      If Not IsNull(VarKbufferHC(1, 0)) Then hasta1 = VarKbufferHC(1, 0)
                      If Not IsNull(VarKbufferHC(2, 0)) Then desde2 = VarKbufferHC(2, 0)
                      If Not IsNull(VarKbufferHC(3, 0)) Then hasta2 = VarKbufferHC(3, 0)
                      If Not IsNull(VarKbufferHC(4, 0)) Then desde3 = VarKbufferHC(4, 0)
                      If Not IsNull(VarKbufferHC(5, 0)) Then hasta3 = VarKbufferHC(5, 0)
                 
                
                If (desde1 - margen) < #1/1/1900# Or desde1 = #1/1/1900# Then
                 desde1 = #1/1/1900 12:00:01 AM#
                 Else: desde1 = desde1 - margen
                End If
                      
                If (hasta1 + margen) > #1/1/1900 11:59:59 PM# Or hasta1 = #1/1/1900# Then
                 hasta1 = #1/1/1900 11:59:00 PM#
                 Else: hasta1 = hasta1 + margen
                End If
                                                                         
                If (desde2 - margen) < #1/1/1900# Or desde2 = #1/1/1900# Then
                 desde2 = #1/1/1900 12:00:01 AM#
                 Else: desde2 = desde2 - margen
                End If
                      
                If (hasta2 + margen) > #1/1/1900 11:59:00 PM# Then
                 hasta2 = #1/1/1900 11:59:00 PM#
                 Else: hasta2 = hasta2 + margen
                End If
                    
                If (desde3 - margen) < #1/1/1900# Then
                 desde3 = #1/1/1900 12:00:01 PM#
                 Else: desde3 = desde3 - margen
                End If
                      
                If (hasta3 + margen) > #1/1/1900 11:59:00 PM# Then
                 hasta3 = #1/1/1900 11:59:00 PM#
                 Else: hasta3 = hasta3 + margen
                End If
                      
                      
                        If (CDate("1/1/1900 " + CStr(HREv)) > desde1 And CDate("1/1/1900 " + CStr(HREv)) < hasta1) Then

                        ElseIf (CDate("1/1/1900 " + CStr(HREv)) > desde2 And CDate("1/1/1900 " + CStr(HREv)) < hasta2) Then

                        ElseIf (CDate("1/1/1900 " + CStr(HREv)) > desde3 And CDate("1/1/1900 " + CStr(HREv)) < hasta3) Then

                        Else
                        Dim HoraI, HoraF As Date
                        HoraI = CDate(HREv) + "00:15"

                        HoraI = Format(HoraI, "hh:mm")

                        HoraF = TimeValue(CDate(HREv)) - TimeValue("00:15")
                        HoraF = Format(HoraF, "hh:mm")


                            'Consultar si esta dentro de 15 min antes/despues y aumentarlo o retrasarlo 15 min
                            'Comprobar si esta dentro del campo restandole 15 min a la hora final y sumandole 15 a la inicial, si es así comprobamos que la hora sea mas grande que el horario de entrega,
                            'si es así le restaremos 15 min para que este dentro de horario, si no le sumaremos 15 min.
                                If (CDate("1/1/1900 " + CStr(HoraI)) > desde1 And CDate("1/1/1900 " + CStr(HoraF)) < hasta1) Then
                                    If (CDate("1/1/1900 " + CStr(HREv)) > hasta1) Then
                                        HREv = HoraF
                                    Else
                                        HREv = HoraI
                                    End If
                                ElseIf (CDate("1/1/1900 " + CStr(HoraI)) > desde2 And CDate("1/1/1900 " + CStr(HoraF)) < hasta2) Then
                                     If (CDate("1/1/1900 " + CStr(HREv)) > hasta2) Then
                                        HREv = HoraF
                                    Else
                                        HREv = HoraI
                                    End If
                                ElseIf (CDate("1/1/1900 " + CStr(HoraI)) > desde3 And CDate("1/1/1900 " + CStr(HoraF)) < hasta3) Then
                                     If (CDate("1/1/1900 " + CStr(HREv)) > hasta3) Then
                                        HREv = HoraF
                                    Else
                                        HREv = HoraI
                                    End If
                                Else
                                'comprobamos si hemos indicado que queremos parar eventos.
                                 If Not (DS.DA.QueryCommand("SELECT * FROM PARAREVE", rowev, lngev)) Then
                                 End If
                                 If (rowev(0, 0) = 1) And IDPalm <> "0" Then
                                        If Not CrearInsert("GESPARAR", 1) Then
                                        End If
                                                 
                                   ' Desmarcamos el viaje como entregado en Resoil
                                    DS.DA.ExecuteQuery ("UPDATE VIAGVVC0 SET VVCFLST5 = '' WHERE (VVCIDVVT = '" & VarIDViag & "') AND (VVCIDADZ = '" & Cliente & "')")
                                    Motivo = "La entrega esta fuera de horario solicitado."
                                    If Not CrearInsert("GESOBCEP", Motivo) Then
                                    PararEvento = True
                                    End If
                             
                                 End If
                                End If
                        End If
                      End If
                    End If
                    Dim rowparados As Variant
                    Dim lngpara As Long
                    'en el caso de que sea un 006 y se quieran parar eventos, buscamos si existe un evento 005 parado y en el caso de que exista, para el 006 para que no se transmita a cepsa y informe de la hora de llegada
                    If (CodEv = "006") Then
                    
                            If Not (DS.DA.QueryCommand("SELECT * FROM PARAREVE", rowev, lngev)) Then
                             End If
                             If (rowev(0, 0) = 1 And Not IsNull(VarKbufferGes(1, i))) Then
                                    If Not (DS.DA.QueryCommand("SELECT GESNUMPD FROM VIAGGES0 WHERE (GESCODEV = '005') AND (GESPARAR = 1) AND (GESNUMPD = ?) AND (GESNUMPR = ?)", rowparados, lngpara, CStr(VarKbufferGes(1, i)), CStr(VarKbufferGes(2, i)))) Then
                                    End If
                                    If (lngpara > 0) Then
                                        If Not CrearInsert("GESPARAR", 1) Then
                                        End If
                                    Motivo = "La entrega esta fuera de horario solicitado."
                                    If Not CrearInsert("GESOBCEP", Motivo) Then
                                    End If
                                    PararEvento = True
                                    End If
                             End If
                             
                             ' 2/12/2020 Si ya se ha enviado un 006 lo paramos
                             If PararEvento = False Then
                             
                                    If Not (DS.DA.QueryCommand("SELECT GESNUMPD FROM VIAGGES0 WHERE (GESCODEV = '006') AND (GESNUMPD = ?) AND (GESNUMPR = ?)", rowparados, lngpara, CStr(VarKbufferGes(1, i)), CStr(VarKbufferGes(2, i)))) Then
                                    End If
                                    If (lngpara > 0) Then
                                        If Not CrearInsert("GESPARAR", 1) Then
                                        End If
                                        Motivo = "Evento 006 ya enviado anteriormente."
                                        If Not CrearInsert("GESOBCEP", Motivo) Then
                                        End If
                                    PararEvento = True
                                    End If
                           
                             End If
                             
                     End If
            'Ponemos la posición del cliente, en el caso de no tener esta posición buscamos la ultima posición del vehiculo y georreferenciamos el cliente.
            If DS.DA.QueryCommand(SQL_POS_CLI, VarKbufferPOSCLI, lngPOSCLI, Cliente) Then
                 
                      If ((Not IsEmpty(VarKbufferPOSCLI(0, 0)) And (Not IsEmpty(VarKbufferPOSCLI(1, 0)) And (VarKbufferPOSCLI(0, 0) <> 0)) And (VarKbufferPOSCLI(0, 0) <> 0) And (VarKbufferPOSCLI(0, 0) <> "") And (VarKbufferPOSCLI(0, 0) <> ""))) Then
                            CordX = CStr(VarKbufferPOSCLI(0, 0))
                            CordY = CStr(VarKbufferPOSCLI(1, 0))
                            CordX = Round(CordX, 5)
                            CordY = Round(CordY, 5)
                            If Not CrearInsert("GESCORDX", CordX) Then 'CORDENADA
                            End If
                            If Not CrearInsert("GESCORDY", CordY) Then 'CORDENADA
                            End If
                

                     Else
                        If (CordY = 0 Or CordX = 0) Then
                                               
                        ' Posicion tecnocom
'                            If Not DS.DA.QueryCommand("SELECT ANAGAAM0.AAMIDDEU FROM VIAGVVT0 INNER JOIN  ANAGAAM0 ON VIAGVVT0.VVTIDAM1 = ANAGAAM0.AAMIDAAM WHERE     (VIAGVVT0.VVTIDVVT = ?)", Varbuff, 1, VarIDViag) Then
'                            End If
'                            If Not IsEmpty(Varbuff) Then idvehi = Varbuff(0, 0)
'                            If Not DS.DA.QueryCommand("SELECT TOP 1 IDVEHIC, Latitud, Longitud, FEUTCVAL From PosicionesOesia WHERE (IDVEHIC = ?) AND (FEUTCVAL < ?) ORDER BY FEUTCVAL DESC", Varbuff, 1, idvehi, FechaEvComp) Then
'                            End If
                        '            Posicion Click & Find
         
                        If Not DS.DA.QueryCommand("SELECT VVTIDAM1, VVTIDAM2 FROM VIAGVVT0 WHERE (VVTIDVVT = ?)", Varbuff, 1, VarIDViag) Then
                        End If
                        If Not IsEmpty(Varbuff) Then
                            Mattr = Varbuff(0, 0)
                            If Not IsNull(Varbuff(1, 0)) Then
                                If Not IsEmpty(Varbuff(1, 0)) And Varbuff(1, 0) <> "" Then MatCis = Varbuff(1, 0)
                            End If
                        End If
                        
                        Matclick = Mattr
                        
                        If Not IsNull(MatCis) Then
                            If Not IsEmpty(MatCis) And MatCis <> "" Then Matclick = MatCis
                        End If
                        
                        If Not DS.DA.QueryCommand("SELECT TOP (1) MATRICULA, LATITUD, LONGITUD, FECHAVAL FROM PosicionesClick WHERE (MATRICULA = ?) AND (FECHAVAL < ?) ORDER BY FECHAVAL DESC", Varbuff, 1, Matclick, FechaEvComp) Then
                        End If '
                        
                            If Not IsEmpty(Varbuff) Then
                                If Not IsNull(Varbuff(1, 0)) Then CordX = Varbuff(1, 0)
                                If Not IsNull(Varbuff(2, 0)) Then CordY = Varbuff(2, 0)
                                If Not IsNull(Varbuff(3, 0)) Then DatePos = Varbuff(3, 0)
                            End If
                            If DatePos = "" Then DatePos = "01/01/1900 0:00:00"
                            DiffHrPos = DateDiff("n", DatePos, FechaEvComp)
                         End If
                        If Not CrearInsert("GESCORDX", CordX) Then ' CORDENADA
                        End If
                        If Not CrearInsert("GESCORDY", CordY) Then ' CORDENADA
                        End If
                        'Geolocalizamos el cliente
                        sqlq = "UPDATE ANAGADZ0 SET ADZCGPSX = '" + CordX + "' ,ADZCGPSY = '" + CordY + "' ,ADZCOOCA = '" + CStr(DiffHrPos) + "' WHERE ADZIDADZ = '" + Cliente + "' "
                        If Not DS.DA.ExecuteQuery(sqlq) Then
                        End If
                    End If
            End If
                    If Not CrearInsert("GESCARR0", CUCarga) Then 'Coeficiente de Utilización
                    End If
                    If Not CrearInsert("GESCODEV", CodEv) Then ' CODIGOEVENTO
                    End If
                    If Not CrearInsert("GESDYEV0", DYEv) Then ' Dia Evento
                    End If
                    HREv = Format(HREv, "hh:mm")
                    If Not CrearInsert("GESHREV0", HREv) Then ' Hora Evento
                    End If
                    If Not CrearInsert("IDVVT", VarIDViag) Then ' CORDENADA
                    End If
                    
                   'Ahora lo hace hasta el 23 por que el 24 es un campo que no va en VIAGGES0
                 For j = 0 To 23  'UBound(VarKbufferGes, 1)
                        If (IsNull(VarKbufferGes(j, i))) Then
                         Value = ""
                        Else
                         Value = VarKbufferGes(j, i)
                        End If
                        columna = SQL_GES_campos_sql005006(j)
                        
                        ' JG 2017 Solo ponemos distancias en 002 y 005
                        If ((CodEv = "002") Or (CodEv = "006")) And (SQL_GES_campos_sql005006(j) = "GESKMCR0") Then
                        If Not IsNumeric(Value) Then Value = "0"
                        ValorKm = CInt(Value)
                        Value = CStr(Round(ValorKm * 1, 0))
                        Else
                        If (SQL_GES_campos_sql005006(j) = "GESKMCR0") Then Value = "0"
                        End If
                        
                        If Not CrearInsert(columna, Value) Then
                        End If
                    Next j
                    
                    ' Paramos el evento para rellenar los compartimentos
                        If Not CrearInsert("GESTRCEP", 0) Then
                        End If
                    
                    ' Comprobamos la hora del evento 003
                    Dim sql_ex03 As String
                    sql_ex03 = "SELECT * FROM VIAGGES0 WHERE IDVVT = ? AND GESCODEV = '003'"
                    If DS.DA.QueryCommand(sql_ex03, buff03, lng03, VarIDViag) Then
                    End If
                    
                    ' Evaluamos si se ha parado el viaje porque lo han ducho manualmente en cuyo caso lo desmarcamos de resoil.
                    
                    If Parar_Datos = True And PararEvento = False Then
                    
                                If Not CrearInsert("GESPARAR", 1) Then
                                End If
                                DS.DA.ExecuteQuery ("UPDATE VIAGVVC0 SET VVCFLST5 = '' WHERE (VVCIDVVT = '" & VarIDViag & "') AND (VVCIDADZ = '" & Cliente & "')")
                                If Not CrearInsert("GESOBCEP", Motivo) Then
                                End If
                                PararEvento = True
                    
                    End If
                    
                    If PararEvento = False Then
                    
                        If lng03 > 0 Then
                        
                            Dim años As String
                            años = Mid(buff03(36, 0), 1, 4)
                            Dim mess As String
                            mess = Mid(buff03(36, 0), 5, 2)
                            Dim dias As String
                            dias = Mid(buff03(36, 0), 7, 2)
                         
                            FechaSalida = CDate(dias & "/" & mess & "/" & años & " " & buff03(37, 0))
                            ' Si la hora del evento es menos de 10 minutos le añadimos 10 minutos
                            If (Round((FechaEvComp - FechaSalida) * 60 * 24) < 10) And (Round((FechaEvComp - FechaSalida) * 60 * 24) >= 0) Then
                                If i = 0 Then
                                    DYEvN = Format(FechaEvComp + ((10 / 60 / 24) - (FechaEvComp - FechaSalida)), "yyyymmdd")
                                    HREvN = Format(FechaEvComp + ((10 / 60 / 24) - (FechaEvComp - FechaSalida)), "hh:mm:ss")
                                    SQL_GEN_EV = "INSERT INTO EventosMBFlotasManual (IdUsuarioMB,CodigoEvento,FechaEvento,FechaServidor,EntaEnLoc,CodigoServicio, CodigoTarea) VALUES ('999','005','" + DYEvN + " " + HREvN + "',GETDATE (),'0','" & CStr(VarIDViag) & "','" & CStr(IdTarea) & "')"
                                    If DS.DA.ExecuteQuery(SQL_GEN_EV) Then
                                    End If
                                End If
                            Else
                            If (Round((FechaEvComp - FechaSalida) * 60 * 24) < 0) Then
                                If Not CrearInsert("GESPARAR", 1) Then
                                End If
                                DS.DA.ExecuteQuery ("UPDATE VIAGVVC0 SET VVCFLST5 = '' WHERE (VVCIDVVT = '" & VarIDViag & "') AND (VVCIDADZ = '" & Cliente & "')")
                                Motivo = "El Evento 003 es posterior"
                                If Not CrearInsert("GESOBCEP", Motivo) Then
                                End If
                                sqlins = "INSERT INTO VIAGGES0 (" + SQL_GES_campos + ") VALUES (" + SQL_GES_values + ")"
                                DS.DA.ExecuteQuery (sqlins)
                                SQL_GES_values = ""
                                SQL_GES_campos = ""
                            Else
                                sqlins = "INSERT INTO VIAGGES0 (" + SQL_GES_campos + ") VALUES (" + SQL_GES_values + ")"
                                DS.DA.ExecuteQuery (sqlins)
                                SQL_GES_values = ""
                                SQL_GES_campos = ""
                                End If
                            End If
                        Else
                            ' Si no hay evento 003 paramos el evento 003
                            If Not CrearInsert("GESPARAR", 1) Then
                            End If
                            ' Desmarcamos el viaje como entregado en Resoil
                            DS.DA.ExecuteQuery ("UPDATE VIAGVVC0 SET VVCFLST5 = '' WHERE (VVCIDVVT = '" & VarIDViag & "') AND (VVCIDADZ = '" & Cliente & "')")
                            Motivo = "No hay evento 003 en este viaje"
                            If Not CrearInsert("GESOBCEP", Motivo) Then
                            End If
                            sqlins = "INSERT INTO VIAGGES0 (" + SQL_GES_campos + ") VALUES (" + SQL_GES_values + ")"
                            DS.DA.ExecuteQuery (sqlins)
                            SQL_GES_values = ""
                            SQL_GES_campos = ""
                        
                        End If
                    Else
                            sqlins = "INSERT INTO VIAGGES0 (" + SQL_GES_campos + ") VALUES (" + SQL_GES_values + ")"
                            DS.DA.ExecuteQuery (sqlins)
                            SQL_GES_values = ""
                            SQL_GES_campos = ""
                    End If
                    'DESPUES DE UN 006 TENEMOS QUE ENVIAR UN 004 DE EL SIGUIENTE VIAJE
                    If (CodEv = "006") And PararEvento = False Then
                        Dim buffprg As Variant
                        Dim lngprg As Long
                        If DS.DA.QueryCommand("SELECT TOP 1 VVCIDVVC From VIAGVVC0 WHERE (VVCIDADZ = ?) AND (VVCIDVVT = ?)", buffprg, lngprg, Cliente, VarIDViag) Then
                        End If
                        If DS.DA.QueryCommand("SELECT TOP 1 VIAGVVC0.VVCIDADZ FROM  VIAGVVT0 INNER JOIN VIAGVVC0 ON VIAGVVT0.VVTIDVVT = VIAGVVC0.VVCIDVVT WHERE (VIAGVVT0.VVTIDVVT = ?) AND (VIAGVVC0.VVCIDADZ <> ?) AND (VIAGVVC0.VVCIDVVC > ?) ORDER BY VIAGVVC0.VVCIDVVC", var004, lng004, VarIDViag, Cliente, CStr(buffprg(0, 0))) Then
                        End If
                        If lng004 > 0 Then
                        'generamos el 004
                
                            If Genera004(VarIDViag, CStr(var004(0, 0)), DYEv, HREv, CordX, CordY) Then
                            End If
                        End If
                    End If
                    
                Next i


ElseIf (CodEv = "007") Then


 If DS.DA.QueryCommand(SQL_GES_007, VarKbufferGes, lngGes, Diaevento, IDPalm) Then
    If lngGes > 0 Then
    VarIDViag = VarKbufferGes(0, 0)
    Cliente = VarKbufferGes(1, 0)
    
 If Not DS.DA.QueryCommand(SQL_GES_005006, VarKbufferGes, lngGes, VarIDViag, Cliente) Then
 End If
 'Calculo de Coeficiente de Utilización
 If Not DS.DA.QueryCommand(SQL_GES_CUCARGA, VarKbufferCU, lngGesCU, VarIDViag) Then
 End If
 If (lngGesCU = 0) Then
     CUCarga = "000.000"
 Else
 If (IsNull(VarKbufferCU(0, 0)) Or IsNull(VarKbufferCU(1, 0)) Or VarKbufferCU(0, 0) = 0 Or VarKbufferCU(1, 0) < 1) Then
         CUCarga = "090.55"
     Else
         CUCarga = CStr((VarKbufferCU(0, 0) / VarKbufferCU(1, 0)) * 100)
     
        CUCarga = Format$(CUCarga, "###.##")
        If (CInt(CUCarga) < 100 And CInt(CUCarga) > 10) Then
           CUCarga = "0" + CUCarga
        ElseIf (CInt(CUCarga) < 10) Then
           CUCarga = "00" + CUCarga
        End If
     End If
 End If
                
 If (lngGes = 0) Then
    GoTo NextIteration
 End If
          
    For i = 0 To UBound(VarKbufferGes, 2)
                
'             Posicion Tecnocom
'            If Not DS.DA.QueryCommand("SELECT ANAGAAM0.AAMIDDEU FROM VIAGVVT0 INNER JOIN  ANAGAAM0 ON VIAGVVT0.VVTIDAM1 = ANAGAAM0.AAMIDAAM WHERE     (VIAGVVT0.VVTIDVVT = ?)", Varbuff, 1, VarIDViag) Then
'            End If
'            If Not IsEmpty(Varbuff) Then idvehi = Varbuff(0, 0)
'            If Not DS.DA.QueryCommand("SELECT TOP 1 IDVEHIC, Latitud, Longitud, FEUTCVAL From PosicionesOesia WHERE (IDVEHIC = ?) AND (FEUTCVAL < ?) ORDER BY FEUTCVAL DESC", Varbuff, 1, idvehi, FechaEvComp) Then
'            End If
'            If Not IsEmpty(Varbuff) Then
'                If Not IsNull(Varbuff(1, 0)) Then CORDX = Varbuff(1, 0)
'                If Not IsNull(Varbuff(2, 0)) Then CORDY = Varbuff(2, 0)
'            End If
                 
'            Posicion Click & Find
         
            If Not DS.DA.QueryCommand("SELECT VVTIDAM1, VVTIDAM2 FROM VIAGVVT0 WHERE (VVTIDVVT = ?)", Varbuff, 1, VarIDViag) Then
            End If
            If Not IsEmpty(Varbuff) Then
                Mattr = Varbuff(0, 0)
                If Not IsNull(Varbuff(1, 0)) Then
                    If Not IsEmpty(Varbuff(1, 0)) And Varbuff(1, 0) <> "" Then MatCis = Varbuff(1, 0)
                End If
            End If
            
            Matclick = Mattr
            
            If Not IsNull(MatCis) Then
                If Not IsEmpty(MatCis) And MatCis <> "" Then Matclick = MatCis
            End If
            
            If Not DS.DA.QueryCommand("SELECT TOP (1) MATRICULA, LATITUD, LONGITUD, FECHAVAL FROM PosicionesClick WHERE (MATRICULA = ?) AND (FECHAVAL < ?) ORDER BY FECHAVAL DESC", Varbuff, 1, Matclick, FechaEvComp) Then
            End If '
            
            If Not IsEmpty(Varbuff) Then
                If Not IsNull(Varbuff(1, 0)) Then CordX = Varbuff(1, 0)
                If Not IsNull(Varbuff(2, 0)) Then CordY = Varbuff(2, 0)
            End If
                 
               
                 
            If Not CrearInsert("GESCORDX", CordX) Then ' CORDENADA
            End If
            If Not CrearInsert("GESCORDY", CordY) Then ' CORDENADA
            End If
                    
                    
            If Not CrearInsert("GESCARR0", CUCarga) Then 'Coeficiente de Utilización
            End If
            If Not CrearInsert("GESCODEV", CodEv) Then ' CODIGOEVENTO
            End If
            If Not CrearInsert("GESDYEV0", DYEv) Then ' Dia Evento
            End If
              HREv = Format(HREv, "hh:mm")
            If Not CrearInsert("GESHREV0", HREv) Then ' Hora Evento
            End If
            If Not CrearInsert("IDVVT", VarIDViag) Then ' IDVIAJE
            End If
        
            
            'Ahora lo hace hasta el 23 por que el 24 es un campo que no va en VIAGGES0
            For j = 0 To 23  'UBound(VarKbufferGes, 1)
                If (IsNull(VarKbufferGes(j, i))) Then
                 Value = ""
                Else
                 Value = VarKbufferGes(j, i)
                End If
                columna = SQL_GES_campos_sql005006(j)
                ' JG 2017 Envío Evento 007 capturar distancia VVTKMPRE para nuevas distancias Cepsa
                ' JG 2017 Solo ponemos distancias en 002 y 005
                If ((CodEv = "002") Or (CodEv = "006")) And (SQL_GES_campos_sql005006(j) = "GESKMCR0") Then
                If Not IsNumeric(Value) Then Value = "0"
                ValorKm = CInt(Value)
                Value = CStr(Round(ValorKm * 1, 0))
                Else
                If (SQL_GES_campos_sql005006(j) = "GESKMCR0") Then Value = "0"
                End If
                If Not CrearInsert(columna, Value) Then
                End If
            Next j
            ' Paramos el evento para rellenar compartimentos
             If Not CrearInsert("GESTRCEP", 0) Then
             End If
            'INSERT INTO VIAGES (SQL_CAMPOS) VALUES (SQL_VALUES)
            sqlins = "INSERT INTO VIAGGES0 (" + SQL_GES_campos + ") VALUES (" + SQL_GES_values + ")"
            DS.DA.ExecuteQuery (sqlins)
            SQL_GES_values = ""
            SQL_GES_campos = ""
            
        Next i
   End If
End If
ElseIf (CodEv = "999") Then
               
                   
        If Not DS.DA.QueryCommand(SQL_GES_005006, VarKbufferGes, lngGes, VarIDViag, Cliente) Then
        End If
                
                
        If (lngGes = 0) Then
            GoTo NextIteration
        End If
                
                 
            For i = 0 To UBound(VarKbufferGes, 2)
               
                                                                
                    
                    If Not CrearInsert("GESCODEV", CodEv) Then ' CODIGOEVENTO
                    End If
                    If Not CrearInsert("GESDYEV0", DYEv) Then ' Dia Evento
                    End If
                      HREv = Format(HREv, "hh:mm")
                    If Not CrearInsert("GESHREV0", HREv) Then ' Hora Evento
                    End If
                    If Not CrearInsert("IDVVT", VarIDViag) Then ' CORDENADA
                    End If
                    
                    ' Comprobamos si existe un evento 003 de este viaje.
                    
                    sql_ex03 = "SELECT * FROM VIAGGES0 WHERE IDVVT = ? AND GESCODEV = '003'"
                    If DS.DA.QueryCommand(sql_ex03, buff03, lng03, VarIDViag) Then
                    End If
                    
                   'Ahora lo hace hasta el 23 por que el 24 es un campo que no va en VIAGGES0
                 
                 For j = 0 To 23  'UBound(VarKbufferGes, 1)
                     'If Not (SQL_GES_campos_sql005006(j) = "GESHRR00" Or SQL_GES_campos_sql005006(j) = "GESDYR00") Then
                        If (IsNull(VarKbufferGes(j, i))) Then
                         Value = ""
                        Else
                         Value = VarKbufferGes(j, i)
                        End If
                        columna = SQL_GES_campos_sql005006(j)
                        ' JG 2017 Solo ponemos distancias en 002 y 005
                        If ((CodEv = "002") Or (CodEv = "006")) And (SQL_GES_campos_sql005006(j) = "GESKMCR0") Then
                        If Not IsNumeric(Value) Then Value = "0"
                        ValorKm = CInt(Value)
                        Value = CStr(Round(ValorKm * 1, 0))
                        Else
                        If (SQL_GES_campos_sql005006(j) = "GESKMCR0") Then Value = "0"
                        End If
                        
                     
                        ' si no hay evento 003 no mandamos la hora de carga
                        If lng03 = 0 Then
                         If SQL_GES_campos_sql(j) = "GESDYTR0" Or SQL_GES_campos_sql(j) = "GESHRTR0" Then Value = ""
                        End If
                        
                        If Not CrearInsert(columna, Value) Then
                        End If
                     'End If
                    Next j
                    
                    'INSERT INTO VIAGES (SQL_CAMPOS) VALUES (SQL_VALUES)
                    sqlins = "INSERT INTO VIAGGES0 (" + SQL_GES_campos + ") VALUES (" + SQL_GES_values + ")"
                    DS.DA.ExecuteQuery (sqlins)
                    SQL_GES_values = ""
                    SQL_GES_campos = ""
                    
                Next i

ElseIf (CodEv = "008") Then

                    
            If Not DS.DA.QueryCommand(SQL_GES_008, VarKbufferGes, lngGes, VarIDViag, Orden) Then
            End If
                
        If (lngGes = 0) Then
            GoTo NextIteration
        End If
                
                 
            For i = 0 To UBound(VarKbufferGes, 2)
            
                    If Not CrearInsert("GESCODEV", CodEv) Then ' CODIGOEVENTO
                    End If
                    If Not CrearInsert("GESDYEV0", DYEv) Then ' Dia Evento
                    End If
                    HREv = Format(HREv, "hh:mm")
                    If Not CrearInsert("GESHREV0", HREv) Then ' Hora Evento
                    End If
                    If Not CrearInsert("IDVVT", VarIDViag) Then ' IDViaje
                    End If
                     If (IdTarea = -33) Then
                          If Not CrearInsert("GESPARAR", 1) Then
                          End If
                     End If
                
                    'Ahora lo hace hasta el 23 por que el 24 es un campo que no va en VIAGGES0
                    For j = 0 To 23  'UBound(VarKbufferGes, 1)
                        If (IsNull(VarKbufferGes(j, i))) Then
                         Value = ""
                        Else
                         Value = VarKbufferGes(j, i)
                        End If
                        columna = SQL_GES_campos_sql005006(j)
                        ' JG 2017 Solo ponemos distancias en 002 y 005
                        If ((CodEv = "002") Or (CodEv = "006")) And (SQL_GES_campos_sql005006(j) = "GESKMCR0") Then
                        If Not IsNumeric(Value) Then Value = "0"
                        ValorKm = CInt(Value)
                        Value = CStr(Round(ValorKm * 1, 0))
                        Else
                        If (SQL_GES_campos_sql005006(j) = "GESKMCR0") Then Value = "0"
                        End If
                        If Not CrearInsert(columna, Value) Then
                        End If
                    Next j
                   
                    'Ultima posicion del vehiculo
                    ' Posisicon tecnocom
'                    If Not DS.DA.QueryCommand("SELECT ANAGAAM0.AAMIDDEU FROM VIAGVVT0 INNER JOIN  ANAGAAM0 ON VIAGVVT0.VVTIDAM1 = ANAGAAM0.AAMIDAAM WHERE     (VIAGVVT0.VVTIDVVT = ?)", Varbuff, 1, VarIDViag) Then
'                    End If
'                    If Not IsEmpty(Varbuff) Then idvehi = Varbuff(0, 0)
'                    If Not DS.DA.QueryCommand("SELECT TOP 1 IDVEHIC, Latitud, Longitud, FEUTCVAL From PosicionesOesia WHERE (IDVEHIC = ?) AND (FEUTCVAL < ?) ORDER BY FEUTCVAL DESC", Varbuff, 1, idvehi, FechaEvComp) Then
'                    End If
'                    If Not IsEmpty(Varbuff) Then
'                        If Not IsNull(Varbuff(1, 0)) Then CORDX = Varbuff(1, 0)
'                        If Not IsNull(Varbuff(2, 0)) Then CORDY = Varbuff(2, 0)
'                    End If
                
'            Posicion Click & Find
         
                    If Not DS.DA.QueryCommand("SELECT VVTIDAM1, VVTIDAM2 FROM VIAGVVT0 WHERE (VVTIDVVT = ?)", Varbuff, 1, VarIDViag) Then
                    End If
                    If Not IsEmpty(Varbuff) Then
                        Mattr = Varbuff(0, 0)
                        If Not IsNull(Varbuff(1, 0)) Then
                            If Not IsEmpty(Varbuff(1, 0)) And Varbuff(1, 0) <> "" Then MatCis = Varbuff(1, 0)
                        End If
                    End If
                    
                    Matclick = Mattr
                    
                    If Not IsNull(MatCis) Then
                        If Not IsEmpty(MatCis) And MatCis <> "" Then Matclick = MatCis
                    End If
                    
                    If Not DS.DA.QueryCommand("SELECT TOP (1) MATRICULA, LATITUD, LONGITUD, FECHAVAL FROM PosicionesClick WHERE (MATRICULA = ?) AND (FECHAVAL < ?) ORDER BY FECHAVAL DESC", Varbuff, 1, Matclick, FechaEvComp) Then
                    End If '
                    
                    If Not IsEmpty(Varbuff) Then
                        If Not IsNull(Varbuff(1, 0)) Then CordX = Varbuff(1, 0)
                        If Not IsNull(Varbuff(2, 0)) Then CordY = Varbuff(2, 0)
                    End If
                
                
                    If Not CrearInsert("GESCORDX", CordX) Then ' CORDENADA
                    End If
                    If Not CrearInsert("GESCORDY", CordY) Then ' CORDENADA
                    End If
                    
                    ' Paramos el evento para rellenar compartimentos
                     If Not CrearInsert("GESTRCEP", 0) Then
                     End If
                     
                     ' Si es manual lo paramos tambien
                     If (IdTarea <> -33) Then
                          If Not CrearInsert("GESPARAR", 1) Then
                          End If
                     End If
                    
                    'INSERT INTO VIAGES (SQL_CAMPOS) VALUES (SQL_VALUES)
                    sqlins = "INSERT INTO VIAGGES0 (" + SQL_GES_campos + ") VALUES (" + SQL_GES_values + ")"
                    DS.DA.ExecuteQuery (sqlins)
                    SQL_GES_values = ""
                    SQL_GES_campos = ""
                    
             Next i

'ElseIf (CInt(CodEv) > 8 And CInt(CodEv) < 21) Then
'JG 25/11/21 Nuevos eventos incidencias

ElseIf ((CInt(CodEv) > 8 And CInt(CodEv) < 26)) Or CInt(CodEv) = 34 Then
            If Not DS.DA.QueryCommand(SQL_GES_005006, VarKbufferGes, lngGes, VarIDViag, Cliente) Then
            End If
                
            If (lngGes = 0) Then
                GoTo NextIteration
            End If
                    
                 
            For i = 0 To UBound(VarKbufferGes, 2)
            
                    If Not CrearInsert("GESCODEV", CodEv) Then ' CODIGOEVENTO
                    End If
                    If Not CrearInsert("GESDYEV0", DYEv) Then ' Dia Evento
                    End If
                    HREv = Format(HREv, "hh:mm")
                    If Not CrearInsert("GESHREV0", HREv) Then ' Hora Evento
                    End If
                    If Not CrearInsert("IDVVT", VarIDViag) Then ' IDViaje
                    End If
                    If Not CrearInsert("GESOBSINC", ObservaInc) Then ' Observaciones incidencias JG 25/11/21
                    End If
                
                    'Ahora lo hace hasta el 23 por que el 24 es un campo que no va en VIAGGES0
                    For j = 0 To 23  'UBound(VarKbufferGes, 1)
                        If (IsNull(VarKbufferGes(j, i))) Then
                         Value = ""
                        Else
                         Value = VarKbufferGes(j, i)
                        End If
                        columna = SQL_GES_campos_sql005006(j)
                        ' JG 2017 Solo ponemos distancias en 002 y 005
                        If ((CodEv = "002") Or (CodEv = "006")) And (SQL_GES_campos_sql005006(j) = "GESKMCR0") Then
                        If Not IsNumeric(Value) Then Value = "0"
                        ValorKm = CInt(Value)
                        Value = CStr(Round(ValorKm * 1, 0))
                        Else
                        If (SQL_GES_campos_sql005006(j) = "GESKMCR0") Then Value = "0"
                        End If
                        
                        If Not CrearInsert(columna, Value) Then
                        End If
                    Next j
                   
                    'Ultima posicion del vehiculo
                    ' Posisicion tecnocom
'                    If Not DS.DA.QueryCommand("SELECT ANAGAAM0.AAMIDDEU FROM VIAGVVT0 INNER JOIN  ANAGAAM0 ON VIAGVVT0.VVTIDAM1 = ANAGAAM0.AAMIDAAM WHERE     (VIAGVVT0.VVTIDVVT = ?)", Varbuff, 1, VarIDViag) Then
'                    End If
'                    If Not IsEmpty((Varbuff)) Then idvehi = Varbuff(0, 0)
'                    If Not DS.DA.QueryCommand("SELECT TOP 1 IDVEHIC, Latitud, Longitud, FEUTCVAL From PosicionesOesia WHERE (IDVEHIC = ?) AND (FEUTCVAL < ?) ORDER BY FEUTCVAL DESC", Varbuff, 1, idvehi, FechaEvComp) Then
'                    End If
'                    If Not IsEmpty(Varbuff) Then
'                        If Not IsNull((Varbuff)) Then CORDX = Varbuff(1, 0)
'                        If Not IsNull(Varbuff) Then CORDY = Varbuff(2, 0)
'                    End If
                
'            Posicion Click & Find
         
                    If Not DS.DA.QueryCommand("SELECT VVTIDAM1, VVTIDAM2 FROM VIAGVVT0 WHERE (VVTIDVVT = ?)", Varbuff, 1, VarIDViag) Then
                    End If
                    If Not IsEmpty(Varbuff) Then
                        Mattr = Varbuff(0, 0)
                        If Not IsNull(Varbuff(1, 0)) Then
                            If Not IsEmpty(Varbuff(1, 0)) And Varbuff(1, 0) <> "" Then MatCis = Varbuff(1, 0)
                        End If
                    End If
                    
                    Matclick = Mattr
                    
                    If Not IsNull(MatCis) Then
                        If Not IsEmpty(MatCis) And MatCis <> "" Then Matclick = MatCis
                    End If
                    
                    If Not DS.DA.QueryCommand("SELECT TOP (1) MATRICULA, LATITUD, LONGITUD, FECHAVAL FROM PosicionesClick WHERE (MATRICULA = ?) AND (FECHAVAL < ?) ORDER BY FECHAVAL DESC", Varbuff, 1, Matclick, FechaEvComp) Then
                    End If '
                    
                    If Not IsEmpty(Varbuff) Then
                        If Not IsNull(Varbuff(1, 0)) Then CordX = Varbuff(1, 0)
                        If Not IsNull(Varbuff(2, 0)) Then CordY = Varbuff(2, 0)
                    End If
                
                    If Not CrearInsert("GESCORDX", CordX) Then ' CORDENADA
                    End If
                    If Not CrearInsert("GESCORDY", CordY) Then ' CORDENADA
                    End If
                    
                    'Buscamos los campos de incidencia para introducirlos en el GES
                    Dim sql_inc As String
                    Dim VarbuffINC As Variant
                    
                    sql_inc = " SELECT MotivoINC,CantidadInvolucrada,CodCLHprod1,CodCLHprod2,ResueltoMediosPropios,TiempoParado,DestinoCep  FROM EventosMBFlotasmanual where IdRegistro = ?"
                    If Not DS.DA.QueryCommand(sql_inc, Varbuff, 1, IdRegistro) Then
                    End If
                    
                     'Agregamos las lineas nuevas de Incidencias (SM 16/08/2011)
                     
                     Dim GESRAINC As String
                     Dim GESLITROI As String
                     Dim GESCLHPRO1 As String
                     Dim GESCLHPRO2 As String
                     Dim GESRESUEL As String
                     Dim GESTIMEPAR As String
                     Dim GESCODDES As String
                     Dim VarBuffPR As Variant
                     Dim LngPr As Long
                     Dim ProdBus As String
                      GESRAINC = ""
                      GESLITROI = ""
                      GESCLHPRO1 = ""
                      GESCLHPRO2 = ""
                      GESRESUEL = ""
                      GESTIMEPAR = ""
                      GESCODDES = ""
                                    
                                    
                     If Not IsNull(Varbuff(0, 0)) Then
                         GESRAINC = Varbuff(0, 0)
                     End If
                     If Not IsNull(Varbuff(1, 0)) Then
                         GESLITROI = Varbuff(1, 0)
                     End If
                     If Not IsNull(Varbuff(2, 0)) Then
                            ProdBus = CStr(Varbuff(2, 0))
                            If Len(ProdBus) > 6 Then ProdBus = ""
                            If Not DS.DA.QueryCommand("SELECT SPRCODSO FROM MULTSPR0 WHERE (SPRIDASC = '901') AND (SPRIDAPR = ?)", VarBuffPR, LngPr, ProdBus) Then
                            End If
                            If LngPr > 0 Then
                            GESCLHPRO1 = Left(VarBuffPR(0, 0), 7)
                            End If
                     End If
                     If Not IsNull(Varbuff(3, 0)) Then
                            ProdBus = CStr(Varbuff(3, 0))
                            If Len(ProdBus) > 6 Then ProdBus = ""
                            If Not DS.DA.QueryCommand("SELECT SPRCODSO FROM MULTSPR0 WHERE (SPRIDASC = '901') AND (SPRIDAPR = ?)", VarBuffPR, LngPr, ProdBus) Then
                            End If
                            If LngPr > 0 Then
                            GESCLHPRO2 = Left(VarBuffPR(0, 0), 7)
                            End If
                     End If
                     If Not IsNull(Varbuff(4, 0)) Then
                        If Varbuff(4, 0) = True Then
                            GESRESUEL = "1"
                        Else
                            GESRESUEL = "0"
                        End If
                     End If
                     If Not IsNull(Varbuff(5, 0)) Then
                         GESTIMEPAR = Varbuff(5, 0)
                     End If
                     If Not IsNull(Varbuff(6, 0)) Then
                         GESCODDES = Varbuff(6, 0)
                     End If
                     
                     
                    If Not CrearInsert("GESRAINC", GESRAINC) Then
                    End If
                    If Not CrearInsert("GESLITROI", Trim(GESLITROI)) Then
                    End If
                    If Not CrearInsert("GESCLHPRO1", GESCLHPRO1) Then
                    End If
                    If Not CrearInsert("GESCLHPRO2", GESCLHPRO2) Then
                    End If
                    If Not CrearInsert("GESRESUEL", GESRESUEL) Then
                    End If
                    If Not CrearInsert("GESTIMEPAR", GESTIMEPAR) Then
                    End If
                    If Not CrearInsert("GESCODDES", GESCODDES) Then
                    End If
                    
                   
                    
                    
                    
                    
                    
                    ' Paramos el evento para rellenar compartimentos
                     If Not CrearInsert("GESTRCEP", 0) Then
                     End If
                     
                     
                     
                    ' Paramos el evento especial para que el usuario seleccione desde Resoil3 que pedido enviar
                
                    Select Case CInt(CodEv)

                    Case 11, 12, 13, 16, 17, 18
                    If Not CrearInsert("GESPARAR", 1) Then
                    End If
                    
                    End Select
                    
                    
                    
                    'INSERT INTO VIAGES (SQL_CAMPOS) VALUES (SQL_VALUES)
                    sqlins = "INSERT INTO VIAGGES0 (" + SQL_GES_campos + ") VALUES (" + SQL_GES_values + ")"
                    If Not DS.DA.ExecuteQuery(sqlins) Then
                        'error en la querry
                        'falta por atrapar
                        Dim rrrr As Integer
                        rrrr = 1
                    End If
                    SQL_GES_values = ""
                    SQL_GES_campos = ""
                    
                Next i
End If





NextIteration:

End Function


Private Function Aleatorio(Minimo As Long, Maximo As Long) As Long
    Randomize ' inicializar la semilla
    Aleatorio = CLng((Minimo - Maximo) * Rnd + Maximo)
End Function



Function ScanEvents_SMS(ByRef LastID As Variant) As Boolean
    Dim i As Long
    Dim j As Long
    Dim varKBuffer As Variant
    Dim varBuffer As Variant
    Dim varKBuffer2 As Variant
    Dim varBufferConsegna As Variant
    Dim varDataEvento As Variant
    Dim varOraEvento As Variant
    Dim varTelAutista As Variant
    Dim strLastDate As String
    Dim lngHowMany As Long
    Dim varIDViaggio As Variant
    Dim varIDConsegna As Variant
    Dim varIDProdotto As Variant
    Dim varIDAutista As Variant
    Dim varIDSMS As Variant
    Dim varTestoSMS As Variant
    Dim VVP As CViaggioProdotto
    Dim VVT As CViaggio
    Dim VVC As CViaggioConsegna
    Dim AAU As CAutista
    Dim dateLocalEventDateTime As Date
    Dim sngKm As Single
    Dim strSQLStatement As String
    Dim strTipoEvento As String
    Dim blnUsaDatiGPS As Boolean
    Dim SQL_INSERT_PALM As String
    
    Dim intMinRitardoInizioViaggio As Integer
'    Dim strSQL_UPDATE_ORA_PREVISTA_CONSEGNE As String
    
    Dim varBufDuoble As Variant
    
    Dim varNRSEQ As Variant
    
    objLogger.logDebug "Inizio ScanEvents_SMS."
    
    varIDSMS = LastID
    
    lngHowMany = 0
    If Not DS.DA.QueryCommand(SQL_SMS_GET_EVENTS, varKBuffer, lngHowMany, varIDSMS) Then
        objLogger.logError "Query SQL_SMS_GET_EVENTS."
        GoTo ErrHandler
    End If

    Set VVT = New CViaggio
    With VVT
        Set .DS.DA = DS.DA
        Set .DS.ErrH = DS.DA.ErrH
    End With
    
    Set VVP = New CViaggioProdotto
    With VVP
        Set .DS.DA = DS.DA
        Set .DS.ErrH = DS.DA.ErrH
    End With
    
    Set VVC = New CViaggioConsegna
    With VVC
        Set .DS.DA = DS.DA
        Set .DS.ErrH = DS.DA.ErrH
    End With
    
    Set AAU = New CAutista
    With AAU
        Set .DS.DA = DS.DA
        Set .DS.ErrH = DS.DA.ErrH
    End With
    
    objLogger.logInfo "Eventi trovati: " & lngHowMany
    
    If lngHowMany > 0 Then
        For i = LBound(varKBuffer, 2) To UBound(varKBuffer, 2)
            ' Inizia transazione
            DS.DA.BeginTransaction
            
            varIDViaggio = ""
            varIDConsegna = ""
            varIDProdotto = ""
            strTipoEvento = vbNullString
            
            objLogger.logInfo "Evento n. " & i + 1
            objLogger.logInfo "-------------------------------------------------"
            objLogger.logInfo "IDSMS         [" & varKBuffer(0, i) & "]"
            objLogger.logInfo "NRTEL         [" & varKBuffer(1, i) & "]"
            objLogger.logInfo "DATASMS       [" & varKBuffer(2, i) & "]"
            objLogger.logInfo "DATALOC       [" & varKBuffer(3, i) & "]"
            objLogger.logInfo "TESTO         [" & varKBuffer(4, i) & "]"
            objLogger.logInfo "-------------------------------------------------"
            
            '-- DP 20/06/08
            '-- Controllo se esiste già un altro SMS per lo stesso autista e con lo stesso numero nella stessa data
            If DS.DA.QueryCommand(SQL_FIND_DUBLE_SMS, varBufDuoble, 0, CVar(varKBuffer(0, i)), CVar(varKBuffer(1, i)), CVar(varKBuffer(2, i)), CVar(DateAdd("d", -1, CDate(varKBuffer(2, i)))), CVar(varKBuffer(4, i))) Then
                If Not IsEmpty(varBufDuoble) Then
                    objLogger.logWarning "SMS scartato in quanto doppio del messaggio " & varBufDuoble(0, 0)
                    DS.DA.EndTrans etRollback
                    GoTo NextIteration
                End If
            Else
                objLogger.logWarning "Errore durante il controllo doppi"
                DS.DA.EndTrans etRollback
                GoTo NextIteration
            End If

            varIDSMS = varKBuffer(0, i)
            
            ' Cerchiamo l'autista corrispondente al numero di telefono del messaggio che stiamo elaborando
            varTelAutista = varKBuffer(1, i)
            If Len(varTelAutista) = 6 And Left(varTelAutista, 1) = "+" Then
                varTelAutista = Mid(varTelAutista, 2)
            ElseIf Len(varTelAutista) > 6 And Left(varTelAutista, 3) = "+34" Then
                varTelAutista = Mid(varTelAutista, 4)
            End If
            
            varIDAutista = AAU.FindByNrTelefono(varTelAutista)
            If IsEmpty(varIDAutista) Then
                objLogger.logWarning "Numero di telefono " & varTelAutista & " non associato ad alcun Autista."
                DS.DA.EndTrans etRollback
                GoTo NextIteration
                'GoTo ErrHandler
            End If
            
            varDataEvento = Format(varKBuffer(2, i), "dd/mm/yyyy")
            varOraEvento = Format(varKBuffer(2, i), "hh:nn")
            varTestoSMS = varKBuffer(4, i)
            varTestoSMS = Replace(varTestoSMS, ".", "")
            varTestoSMS = Replace(varTestoSMS, ",", "")
            varTestoSMS = Trim(varTestoSMS)
            If Not IsNumeric(varTestoSMS) Then
                objLogger.logError "Testo SMS non numerico"
                DS.DA.EndTrans etRollback
                GoTo NextIteration
                'GoTo ErrHandler
            End If
            
            varIDViaggio = VVT.FindBySMSCarico(varIDAutista, varDataEvento, varTestoSMS)
            
            ' Determiniamo se si tratta di un evento di fine carico o di inizio scarico
            Select Case varIDViaggio
                Case "0"
                '005
                    ' Non è un carico, cerchiamo tra gli scarichi
                    varBufferConsegna = VVC.FindBySMSConsegna(varIDAutista, varDataEvento, varTestoSMS)
                    If IsEmpty(varBufferConsegna) Then
                        objLogger.logWarning "Errore durante la ricerca della Consegna, Autista " & varIDAutista & ", Data " & varDataEvento
                        DS.DA.EndTrans etRollback
                        GoTo NextIteration
                        'GoTo ErrHandler
                    Else
                        varIDViaggio = varBufferConsegna(0, 0)
                        varIDConsegna = varBufferConsegna(2, 0)
                        '--SM Inseción de los SMS en el flujo de PDA.
                        SQL_INSERT_PALM = "INSERT INTO EventosMBFlotasManual (IdTareaMB,IdUsuarioMB,CodigoServicio,CodigoTarea,CodigoEvento,FechaEvento,FechaServidor,EntaEnLoc) VALUES ('666666','666','" + varIDViaggio + "','" + varIDConsegna + "','005','" + varDataEvento + " " + varOraEvento + "',GETDATE (),'0')"
                        If DS.DA.ExecuteQuery(SQL_INSERT_PALM) Then
                            'EVENTO GENERADO EN PALMS
                        End If
                        
                        strTipoEvento = "INI"
                    End If
                    
                Case "-1"
                    objLogger.logError "Errore durante la ricerca del Viaggio, Autista " & varIDAutista & ", Data " & varDataEvento
                    DS.DA.EndTrans etRollback
                    GoTo NextIteration
                    'GoTo ErrHandler
                    
                Case Else
                '003
                    ' Si tratta di un carico

                        '--SM Inseción de los SMS en el flujo de PDA.
                        SQL_INSERT_PALM = "INSERT INTO EventosMBFlotasManual (IdTareaMB,IdUsuarioMB,CodigoServicio,CodigoEvento,FechaEvento, FechaServidor,EntaEnLoc) VALUES ('666666','666','" + varIDViaggio + "','003','" + varDataEvento + " " + varOraEvento + "',GETDATE (),'0')"
                        If DS.DA.ExecuteQuery(SQL_INSERT_PALM) Then
                            'EVENTO GENERADO EN PALMS
                        End If
                                     
                        
                    varIDConsegna = vbNullString
                    varIDProdotto = vbNullString
                    strTipoEvento = "SAL"
                    
            End Select
            
'            objLogger.logInfo "RESOil Keys VVTIDVVT = """ & varIDViaggio & _
'                                """, VVCIDVVC = """ & varIDConsegna & """"
'
'            ' Se non trovo il viaggio, skip
'            If varIDViaggio <> "" Then
'                ' Creazione dell'evento
'                DS.Fields("0").NewValue = DS.DA.GetCounter("anaganu0", "anuprogr", "anuidanu", "904", 0)
'                DS.Fields("0").Updated = True
'
'                DS.Fields("1").NewValue = varIDViaggio
'                DS.Fields("1").Updated = True
'
'                DS.Fields("2").NewValue = varIDConsegna
'                DS.Fields("2").Updated = True
'
'                DS.Fields("3").NewValue = varIDProdotto
'                DS.Fields("3").Updated = True
'
'                DS.Fields("5").NewValue = varDataEvento
'                DS.Fields("5").Updated = True
'
'                DS.Fields("9").NewValue = "K"
'                DS.Fields("9").Updated = True
'
'                ' Carica il viaggio
'                With VVT
'                    .DS.Fields.ResetFields
'                    .DS.Fields("0").FindValue = varIDViaggio
'
'                    If Not .Find Then
'                        objLogger.logError "Find Viaggio."
'                        DS.DA.EndTrans etRollback
'                        GoTo NextIteration
'                        'GoTo ErrHandler
'                    End If
'
'                    blnUsaDatiGPS = True
'                End With
'
'
'                Select Case strTipoEvento
'                    ' -------------------------------------------------------
'                    ' Evento - Salida de factora (Fine del carico)
'                    ' -------------------------------------------------------
'                    Case "SAL"
'                        DS.Fields("4").NewValue = "CF"
'                        DS.Fields("4").Updated = True
'
'                        objLogger.logInfo "Tipo evento: Fine carico"
'                        If Not Insert Then
'                            objLogger.logError "Insert ResOil evento."
'                            DS.DA.EndTrans etRollback
'                            GoTo NextIteration
'                            'GoTo ErrHandler
'                        End If
'
'                        ' Aggiorna lo stato del viaggio (VVTFLSTA)
'                        VVT.DS.Fields("48").NewValue = VIAGGIO_STATO_INCORSO
'                        VVT.DS.Fields("48").Updated = True
'                        LogFieldValue VVT.DS.Fields("48")
'
'                        ' Aggiorna il tempo di carico sul viaggio (VVTTMPCA)
'                        If blnUsaDatiGPS Then
'                            ' -- DP 06/05/08
'                            ' -- Oltre alle ore aggiorno anche le date visto che alcuni carichi avvengono a cavallo della mezzanotte
'
'                            ' Aggiorna data e ora di inizio viaggio e data ed ora di inizio carico sulla testata viaggio (VVTDTINV, VVTHRINV, VVTDTINC, VVTHRINC)
'                            VVT.DS.Fields("6").NewValue = varDataEvento
'                            VVT.DS.Fields("6").Updated = True
'                            LogFieldValue VVT.DS.Fields("6")
'                            VVT.DS.Fields("7").NewValue = varOraEvento
'                            VVT.DS.Fields("7").Updated = True
'                            LogFieldValue VVT.DS.Fields("7")
'                            VVT.DS.Fields("4").NewValue = varDataEvento
'                            VVT.DS.Fields("4").Updated = True
'                            LogFieldValue VVT.DS.Fields("4")
'                            VVT.DS.Fields("5").NewValue = varOraEvento
'                            VVT.DS.Fields("5").Updated = True
'                            LogFieldValue VVT.DS.Fields("5")
'
'                            ' Aggiorna flag carico completato
'                            VVT.DS.Fields("39").NewValue = "S"
'                            VVT.DS.Fields("39").Updated = True
'                            LogFieldValue VVT.DS.Fields("39")
'
'                            '-- DP 05/06/08
'                            '-- Aggiorno il campo VVTFLAGG, che contiene l'indicazione dell'ultimo processo che ha aggiornato il record
'                            VVT.DS.Fields("52").NewValue = "ME"
'                            VVT.DS.Fields("52").Updated = True
'                            LogFieldValue VVT.DS.Fields("52")
'
'
'                            '-- DP 20/06/08
'                            '-- Questo controllo serve a bypassare eventuali errori dovuti a viaggi senza ora di inizio carico
'                            If CStr(VVT.DS.Fields("7").Value) <> "01/01/1900" Then
'                                '-- DP 13/05/08
'                                '-- Il confronto lo faccio su VVTDTINV + VVTHRINV e l'orario dell'evento della vista VIAGSMS0
'                                intMinRitardoInizioViaggio = DateDiff("n", CDate(VVT.DS.Fields("6").Value & " " & Format(VVT.DS.Fields("7").Value, "hh.mm")), CDate(varKBuffer(2, i)))
'                                If Not VVC.blnUpdateOrePreviste(intMinRitardoInizioViaggio, varIDViaggio) Then
'                                    objLogger.logError "Errore in aggiornamento orario previsto consegne"
'                                    DS.DA.EndTrans etRollback
'                                    GoTo NextIteration
'                                End If
'
''                                If intMinRitardoInizioViaggio >= 50 Then
''                                    '-- In caso di ritardo compreso tra 50 e 60 minuti imposto comunque un ritardo di 60 minuti
''                                    If intMinRitardoInizioViaggio < 60 Then
''                                        intMinRitardoInizioViaggio = 60
''                                    End If
''
''                                    strSQL_UPDATE_ORA_PREVISTA_CONSEGNE = "UPDATE VIAGVVC0 SET VVCDTPRE = CASE WHEN dateadd(n, " & intMinRitardoInizioViaggio & ", VVCTEMVG) > CONVERT(DATETIME, '1900-01-01 23:59:59', 102) THEN dateadd(d, 1, VVCDTCON) ELSE VVCDTCON END, " & _
''                                                                        "       VVCHRPRE = dateadd(n, " & intMinRitardoInizioViaggio & ", VVCTEMVG), VVCFLAGG = 'ME' " & _
''                                                                        "  WHERE VVCIDVVT = '" & CStr(varIDViaggio) & "' AND (VVCFLST5 = 'N' OR VVCFLST5 IS NULL) AND VVCANNUL = 0"
''
''                                    If Not DS.DA.ExecuteQuery(strSQL_UPDATE_ORA_PREVISTA_CONSEGNE) Then
''                                        objLogger.logError "Errore in aggiornamento orario previsto consegne"
''                                        DS.DA.EndTrans etRollback
''                                        GoTo NextIteration
''                                        'GoTo ErrHandler
''                                    End If
''                                End If
'                            End If
'
'
'                            ' Aggiorna il flag modificato (VVPFLST5) di tutti i prodotti della consegna
'                            If Not DS.DA.ExecuteCommand(SQL_UPDATE_VIAGGIO_MODIFICATO, varIDViaggio) Then
'                                objLogger.logError "Errore in aggiornamento flag Modificato del viaggio."
'                                DS.DA.EndTrans etRollback
'                                GoTo NextIteration
'                                'GoTo ErrHandler
'                            End If
'                        End If
'
'                        ' Aggiorna il viaggio
'                        With VVT
'                            If .CheckViaggioCompletato(varIDViaggio) Then
'                                If Not .SetStato("C") Then
'                                    objLogger.logError "Set Viaggio Completato."
'                                    DS.DA.EndTrans etRollback
'                                    GoTo NextIteration
'                                    'GoTo ErrHandler
'                                End If
'                            End If
'
'                            If blnUsaDatiGPS Then
'                                ' Aggiorna il flag trasmissione viaggio
'                                .DS.Fields("41").NewValue = "S"
'                                .DS.Fields("41").Updated = True
'                                LogFieldValue .DS.Fields("41")
'                            End If
'
'                            ' Salva
'                            If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
'                                objLogger.logError "Save Viaggio."
'                                DS.DA.EndTrans etRollback
'                                GoTo NextIteration
'                                'GoTo ErrHandler
'                            End If
'
'                            .DS.DA.EndTrans etCommit
'                        End With
'
'
'                    ' -------------------------------------------------------
'                    ' Evento - Inicio descarga (Inizio scarico prodotto)
'                    ' -------------------------------------------------------
'                    Case "INI"
'                        '-- DP 24/07/08
'                        '-- Se c'è almeno un prodotto già trasmesso su GES non devo più aggiornare i dati
'                        If Not VVP.blnGESAlreadyTransmitted(CStr(varIDViaggio), CStr(varIDConsegna)) Then
'                            objLogger.logInfo "Tipo evento: Inizio consegna prodotto"
'
'                            ' Carica la consegna e aggiorna i dati
'                            With VVC
'                                .DS.Fields.ResetFields
'                                .DS.Fields("0").FindValue = varIDViaggio
'                                .DS.Fields("1").FindValue = varIDConsegna
'
'                                If Not .Find Then
'                                    objLogger.logError "Find Consegna."
'                                    DS.DA.EndTrans etRollback
'                                    GoTo NextIteration
'                                    'GoTo ErrHandler
'                                End If
'
'                                varNRSEQ = .DS.Fields("2").Value
'
'                                If blnUsaDatiGPS Then
'                                    '-- DP 15/05/08
'                                    '-- Aggiorno data e ora di arrivo sui nuovi campi VVCDTARR e VVCHRARR anziché VVCTEMVG
'                                    .DS.Fields("49").NewValue = varDataEvento
'                                    .DS.Fields("49").Updated = True
'                                    LogFieldValue .DS.Fields("49")
'                                    .DS.Fields("47").NewValue = varOraEvento
'                                    .DS.Fields("47").Updated = True
'                                    LogFieldValue .DS.Fields("47")
'
'                                    ' Aggiorna il flag consegna ultimata
'                                    .DS.Fields("36").NewValue = "S"
'                                    .DS.Fields("36").Updated = True
'                                    LogFieldValue .DS.Fields("36")
'
'                                    ' Aggiorna lo stato 'Completo'
'                                    .DS.Fields("44").NewValue = "2"
'                                    .DS.Fields("44").Updated = True
'                                    LogFieldValue .DS.Fields("44")
'
'                                    '-- DP 05/06/08
'                                    '-- Aggiorno il campo VVCFLAGG, che contiene l'indicazione dell'ultimo processo che ha aggiornato il record
'                                    .DS.Fields("50").NewValue = "ME"
'                                    .DS.Fields("50").Updated = True
'                                    LogFieldValue .DS.Fields("50")
'                                End If
'
'                                ' Salva le modifiche sulla consegna
'                                If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
'                                    objLogger.logError "Save Consegna."
'                                    DS.DA.EndTrans etRollback
'                                    GoTo NextIteration
'                                    'GoTo ErrHandler
'                                End If
'                            End With
'
'
'                            '-- DP 20/06/08
'                            '-- Questo controllo serve a bypassare eventuali errori dovuti a consegne senza data o ora (VVTDTCON e VVTTEMVG)
'                            If CStr(VVC.DS.Fields("41").Value) <> "01/01/1900" And CStr(VVC.DS.Fields("24").Value) <> "01/01/1900" Then
'                                '-- DP 30/09/08
'                                '-- L'aggiornamento deve essere fatto solamente se le consegne avvengono nell'ordine previsto
'                                If blnConsegnaInOrdine(varIDViaggio, varNRSEQ) Then
'                                    '-- DP 19/05/08
'                                    '-- Il confronto lo faccio su VVTDTCON + VVTTEMVG e l'orario dell'evento della vista VIAGSMS0
'                                    intMinRitardoInizioViaggio = DateDiff("n", CDate(VVC.DS.Fields("41").Value & " " & Format(VVC.DS.Fields("24").Value, "hh.mm")), CDate(varKBuffer(2, i)))
'                                    If Not VVC.blnUpdateOrePreviste(intMinRitardoInizioViaggio, varIDViaggio) Then
'                                        objLogger.logError "Errore in aggiornamento orario previsto consegne"
'                                        DS.DA.EndTrans etRollback
'                                        GoTo NextIteration
'                                    End If
'
''                                    If intMinRitardoInizioViaggio >= 50 Then
''                                        '-- In caso di ritardo compreso tra 50 e 60 minuti imposto comunque un ritardo di 60 minuti
''                                        If intMinRitardoInizioViaggio < 60 Then
''                                            intMinRitardoInizioViaggio = 60
''                                        End If
''
''                                        strSQL_UPDATE_ORA_PREVISTA_CONSEGNE = "UPDATE VIAGVVC0 SET VVCDTPRE = CASE WHEN dateadd(n, " & intMinRitardoInizioViaggio & ", VVCTEMVG) > CONVERT(DATETIME, '1900-01-01 23:59:59', 102) THEN dateadd(d, 1, VVCDTCON) ELSE VVCDTCON END, " & _
''                                                                            "       VVCHRPRE = dateadd(n, " & intMinRitardoInizioViaggio & ", VVCTEMVG), VVCFLAGG = 'ME' " & _
''                                                                            "  WHERE VVCIDVVT = '" & CStr(varIDViaggio) & "' AND (VVCFLST5 = 'N' OR VVCFLST5 IS NULL) AND VVCANNUL = 0"
''
''                                        If Not DS.DA.ExecuteQuery(strSQL_UPDATE_ORA_PREVISTA_CONSEGNE) Then
''                                            objLogger.logError "Errore in aggiornamento orario previsto consegne"
''                                            DS.DA.EndTrans etRollback
''                                            GoTo NextIteration
''                                            'GoTo ErrHandler
''                                        End If
''                                    End If
'                                End If
'                            End If
'
'
'                            ' Ricerco tutti i prodotti con VVPFLSTA  <> '2'
'                            lngHowMany = 0
'                            If Not DS.DA.QueryCommand(SQL_KNOSOS_GET_COUNT_SCARICHI_NONCONCLUSI_VVT, varBuffer, lngHowMany, varIDViaggio) Then
'                                objLogger.logError "Query SQL_KNOSOS_GET_COUNT_SCARICHI_NONCONCLUSI_VVT."
'                                DS.DA.EndTrans etRollback
'                                GoTo NextIteration
'                                'GoTo ErrHandler
'                            End If
'
'                            VVP.DS.Field("Numero Viaggio").FindValue = varIDViaggio
'                            VVP.DS.Field("Numero Consegna").FindValue = varIDConsegna
'                            If Not VVP.GetList(True) Then
'                                DS.DA.EndTrans etRollback
'                                GoTo NextIteration
'                                'GoTo ErrHandler
'                            End If
'
'                            If varBuffer(0, 0) = 0 Then
'                                ' Tutti i prodotti del viaggio hanno VVPFLSTA  = '2', impostiamo il flag modificato (VVPFLST5) di tutti i prodotti del viaggio
'                                For j = 0 To VVP.DS.HowMany - 1
'                                    ' Carica i prodotti della consegna e aggiorna i dati
'                                    With VVP
'                                        .DS.Fields.ResetFields
'                                        .DS.Fields("0").FindValue = varIDViaggio
'                                        .DS.Fields("1").FindValue = varIDConsegna
'                                        .DS.Fields("2").FindValue = varIDProdotto
'
'                                        If Not .Find Then
'                                            objLogger.logError "Find Consegna Prodotto."
'                                            DS.DA.EndTrans etRollback
'                                            GoTo NextIteration
'                                            'GoTo ErrHandler
'                                        End If
'                                        If .DS.HowMany = 0 Then
'                                            objLogger.logError "Prodotto Consegna non trovato"
'                                            DS.DA.EndTrans etRollback
'                                            GoTo NextIteration
'                                            'GoTo ErrHandler
'                                        End If
'                                    End With
'                                Next j
'
'                                If Not DS.DA.ExecuteCommand(SQL_UPDATE_VIAGGIO_MODIFICATO, varIDViaggio) Then
'                                    objLogger.logError "Errore in aggiornamento flag Modificato del viaggio."
'                                    DS.DA.EndTrans etRollback
'                                    GoTo NextIteration
'                                    'GoTo ErrHandler
'                                End If
'                            Else
'                                For j = 0 To VVP.DS.HowMany - 1
'                                    varIDProdotto = VVP.DS.ListItem(0, j)
'
'                                    DS.Fields("0").NewValue = DS.DA.GetCounter("anaganu0", "anuprogr", "anuidanu", "904", 0)
'                                    DS.Fields("0").Updated = True
'
'                                    DS.Fields("1").NewValue = varIDViaggio
'                                    DS.Fields("1").Updated = True
'
'                                    DS.Fields("2").NewValue = varIDConsegna
'                                    DS.Fields("2").Updated = True
'
'                                    DS.Fields("3").NewValue = varIDProdotto
'                                    DS.Fields("3").Updated = True
'
'                                    DS.Fields("4").NewValue = "SI"
'                                    DS.Fields("4").Updated = True
'
'                                    DS.Fields("5").NewValue = varDataEvento
'                                    DS.Fields("5").Updated = True
'
'                                    DS.Fields("9").NewValue = "K"
'                                    DS.Fields("9").Updated = True
'
'                                    objLogger.logInfo "Creazione dell'evento in ResOil."
'                                    If Not Insert Then
'                                        objLogger.logError "Insert ResOil evento."
'                                        DS.DA.EndTrans etRollback
'                                        GoTo NextIteration
'                                        'GoTo ErrHandler
'                                    End If
'
'                                    ' Carica i prodotti della consegna e aggiorna i dati
'                                    With VVP
'                                        .DS.Fields.ResetFields
'                                        .DS.Fields("0").FindValue = varIDViaggio
'                                        .DS.Fields("1").FindValue = varIDConsegna
'                                        .DS.Fields("2").FindValue = varIDProdotto
'
'                                        If Not .Find Then
'                                            objLogger.logError "Find Consegna Prodotto."
'                                            DS.DA.EndTrans etRollback
'                                            GoTo NextIteration
'                                            'GoTo ErrHandler
'                                        End If
'                                        If .DS.HowMany = 0 Then
'                                            objLogger.logError "Prodotto Consegna non trovato"
'                                            DS.DA.EndTrans etRollback
'                                            GoTo NextIteration
'                                            'GoTo ErrHandler
'                                        End If
'
'                                        ' Aggiorna lo stato dello scarico (VVPFLSTA)
'                                        .DS.Fields("37").NewValue = "2"
'                                        .DS.Fields("37").Updated = True
'                                        LogFieldValue .DS.Fields("37")
'
'                                        ' Aggiorna il flag Modificato (VVPFLST5)
'                                        .DS.Field("Stato Riga 5").NewValue = "S"
'                                        .DS.Field("Stato Riga 5").Updated = True
'                                        LogFieldValue .DS.Field("Stato Riga 5")
'
'                                        '-- DP 05/06/08
'                                        '-- Aggiorno il campo VVPFLAGG, che contiene l'indicazione dell'ultimo processo che ha aggiornato il record
'                                        .DS.Fields("43").NewValue = "ME"
'                                        .DS.Fields("43").Updated = True
'                                        LogFieldValue .DS.Fields("43")
'
'                                        ' Aggiorna l'ora inizio scarico (VVPHRINC)
'                                        If blnUsaDatiGPS Then
'                                            .DS.Fields("38").NewValue = varOraEvento
'                                            .DS.Fields("38").Updated = True
'                                            LogFieldValue .DS.Fields("38")
'                                        End If
'
'                                        ' Salva
'                                        If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
'                                            objLogger.logError "Save Consegna Prodotto."
'                                            DS.DA.EndTrans etRollback
'                                            GoTo NextIteration
'                                            'GoTo ErrHandler
'                                        End If
'                                    End With
'                                Next j
'                            End If
'
'                            ' Aggiorna il viaggio
'                            With VVT
'                                If .CheckViaggioCompletato(varIDViaggio) Then
'                                    If Not .SetStato("C") Then
'                                        objLogger.logError "Set Viaggio Completato."
'                                        DS.DA.EndTrans etRollback
'                                        GoTo NextIteration
'                                        'GoTo ErrHandler
'                                    End If
'                                End If
'
'                                If blnUsaDatiGPS Then
'                                    ' Aggiorna il flag trasmissione viaggio
'                                    .DS.Fields("41").NewValue = "S"
'                                    .DS.Fields("41").Updated = True
'                                    LogFieldValue .DS.Fields("41")
'
'                                    '-- DP 05/06/08
'                                    '-- Aggiorno il campo VVTFLAGG, che contiene l'indicazione dell'ultimo processo che ha aggiornato il record
'                                    .DS.Fields("52").NewValue = "ME"
'                                    .DS.Fields("52").Updated = True
'                                    LogFieldValue .DS.Fields("52")
'                                End If
'
'                                ' Salva
'                                If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
'                                    objLogger.logError "Save Viaggio."
'                                    DS.DA.EndTrans etRollback
'                                    GoTo NextIteration
'                                    'GoTo ErrHandler
'                                End If
'
'                                .DS.DA.EndTrans etCommit
'                            End With
'                        Else
'                            objLogger.logError "Ora di reale di consegna già trasmessa su GES"
'                        End If
'
'                        DS.DA.EndTrans etCommit
'
'                    Case Else
'                        objLogger.logError "Tipo evento sconosciuto."
'                        DS.DA.EndTrans etRollback
'                        GoTo NextIteration
'                        'GoTo ErrHandler
'
'                End Select
'            Else
'                DS.DA.LogExecTime "WARN: Viaggio non trovato!", -1
'            End If
            
            DS.DA.EndTrans etCommit
NextIteration:
        Next i
    End If

    ScanEvents_SMS = True
    
    Set VVP = Nothing
    Set VVT = Nothing
    Set VVC = Nothing

    objLogger.logDebug "Fine ScanEvents_SMS senza errori."
    GoTo EndProc

ErrHandler:
    DS.DA.EndTrans etRollback
    
    Set VVP = Nothing
    Set VVT = Nothing
    Set VVC = Nothing
    
    objLogger.logError "Fine ScanEvents_SMS con ERRORI."

    ScanEvents_SMS = False

EndProc:
    LastID = varIDSMS
    

    
End Function


'Public Function AggiornaOrari_SMS(Data As Variant, OraAttuale As Variant) As Boolean
'    On Error GoTo ErrHandler
'
'    Dim i As Long
'    Dim j As Long
'    Dim k As Long
'    Dim dblMinutiDelay As Double
'    Dim varBuffer As Variant
'    Dim varConsegnaEffettuata As Variant
'    Dim varCaricoEffettuato As Variant
'    Dim varOraPrevistaCarico As Variant
'    Dim varOraPrevistaConsegna As Variant
'    Dim lngHowMany As Long
'    Dim lngBufferCount As Long
'    Dim lngPointer As Long
'    Dim strTargaMotrice As String
'    Dim strIDViaggio As String
'    Dim strIDConsegna As String
'    Dim VVT As CViaggio
'    Dim VVC As CViaggioConsegna
'
''    dblMinutiDelay = 10
'    dblMinutiDelay = 60
'
'    Set VVT = New CViaggio
'    With VVT
'        Set .DS.DA = DS.DA
'        Set .DS.ErrH = DS.DA.ErrH
'    End With
'    Set VVC = New CViaggioConsegna
'    With VVC
'        Set .DS.DA = DS.DA
'        Set .DS.ErrH = DS.DA.ErrH
'    End With
'
'    lngHowMany = 0
'    If Not IsDate(Data) Then
'        GoTo ErrHandler
'    End If
'
'    '-- DP 13/05/08
'    '-- La seguente query ricerca i dati dei viaggi con VVTDTINV = [Data passata alla funzione].
'    '-- Eventuali viaggi a cavallo della mezzanotte potrebbero non essere controllati correttamente
'    Data = Format(Data, "dd/mm/yyyy")
'    If Not DS.DA.QueryCommand(SQL_SMS_GET_LIST_AGG, varBuffer, lngHowMany, Data) Then
'        objLogger.logError "Query SQL_SMS_GET_LIST_AGG."
'        GoTo ErrHandler
'    End If
'    If Not IsEmpty(varBuffer) Then
'        OraAttuale = Format(OraAttuale, "hh:nn")
'        If Not IsDate(OraAttuale) Then
'            GoTo ErrHandler
'        End If
'        lngBufferCount = UBound(varBuffer, 2)
'
''        Dim stra As String
''        For i = 0 To 50
''            stra = Format(i, "00") & " "
''            For j = 0 To 7
''                stra = stra & varBuffer(j, i) & vbTab
''            Next j
''            Debug.Print stra
''        Next i
'
'        For i = 0 To lngBufferCount
'            strTargaMotrice = varBuffer(0, i)
'
'            '-- DP 13/05/08
'            '-- Il seguente ciclo posiziona "lngPointer" sull'ULTIMA consegna EFFETTUATA dell'ULTIMO viaggio
'            '-- del veicolo corrente. Questa logica è corretta solo idealmente, nella realtà, nel caso in cui
'            '-- le consegne fossero invertite, potrei "saltare" il controllo su alcune consegne.
'
'            ' Mi posiziono sulla prima consegna non effettuata
'            lngPointer = i - 1
'            For j = i To lngBufferCount
'                If strTargaMotrice <> varBuffer(0, j) Then
'                    Exit For
'                End If
'                varCaricoEffettuato = varBuffer(3, j)
'                varConsegnaEffettuata = varBuffer(4, j)
'                If (varConsegnaEffettuata = "S") Then
'                    lngPointer = j
'                End If
'            Next j
''            lngPointer = j
'
'            ' Elaboro tutte le consegne non effettuate per la motrice in corso di elaborazione
'            If lngPointer <= lngBufferCount Then
'
'                If varConsegnaEffettuata = "S" Then
'                Else
'
'                    '-- DP 13/05/08
'                    '-- Questo (lngPointer + 1) è corretto solo idealmente come spiegato nel precedente commento
'                    For j = (lngPointer + 1) To lngBufferCount
'                        If strTargaMotrice <> varBuffer(0, j) Then
'                            Exit For
'                        End If
'                        varOraPrevistaCarico = Format(varBuffer(5, j), "hh:nn")
'                        varOraPrevistaConsegna = Format(varBuffer(6, j), "hh:nn")
'
'                        ' Individuo le consegne in ritardo di almeno 50 minuti rispetto all'orario passato come parametro
''                        If (DateDiff("n", varOraPrevistaConsegna, OraAttuale) > 0) Or (DateDiff("n", varOraPrevistaCarico, OraAttuale) > 0) Then
'                        If (DateDiff("n", varOraPrevistaConsegna, OraAttuale) >= 50) Or (DateDiff("n", varOraPrevistaCarico, OraAttuale) >= 50) Then
'
'                            '-- DP 13/05/08
'                            '-- Questo ciclo è inutile, potrebbe bastare quello precedente
'                            For k = j To lngBufferCount
'                                If strTargaMotrice <> varBuffer(0, k) Then
'                                    Exit For
'                                End If
'                                strIDViaggio = varBuffer(1, k)
'                                strIDConsegna = varBuffer(2, k)
'
'                                DS.DA.BeginTransaction
''                                If IsNull(varCaricoEffettuato) Or (varCaricoEffettuato <> "S") Then
'                                If IsNull(varBuffer(3, k)) Or (varBuffer(3, k) <> "S") Then
'                                    varOraPrevistaCarico = Format(varBuffer(5, k), "hh:nn")
'
'                                    '-- DP 13/05/08
'                                    '-- La seguente operazione è un errore, "varOraPrevistaCarico" andrebbe incrementata di "dblMinutiDelay"
'                                    '-- dopo l'"If", non prima, esattamente come avviene per le consegne (più in basso)
'                                    varOraPrevistaCarico = Format(DateAdd("n", dblMinutiDelay, varOraPrevistaCarico), "hh:nn")
'
''                                    If DateDiff("n", varOraPrevistaCarico, OraAttuale) > 0 Then
'                                    If DateDiff("n", varOraPrevistaCarico, OraAttuale) >= 50 Then
'                                        varOraPrevistaCarico = OraAttuale
'                                        Debug.Print "C" & " " & strTargaMotrice & " " & strIDViaggio & " " & strIDConsegna & " " & varOraPrevistaCarico
'
'                                        With VVT
'                                            .DS.Fields.ResetFields
'                                            .DS.Fields("0").FindValue = strIDViaggio
'
'                                            If Not .Find Then
'                                                objLogger.logError "Find Viaggio."
'                                                GoTo ErrHandler
'                                            End If
'                                            .DS.Field("Ora inizio carico").NewValue = varOraPrevistaCarico
'                                            .DS.Field("Ora inizio carico").Updated = True
'                                            If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
'                                                objLogger.logError "Save Viaggio."
'                                                GoTo ErrHandler
'                                            End If
'
'                                            .DS.DA.EndTrans etCommit
'                                        End With
'                                    End If
'                                End If
'
'                                varOraPrevistaConsegna = Format(varBuffer(6, k), "hh:nn")
''                                varOraPrevistaConsegna = Format(DateAdd("n", dblMinutiDelay, varOraPrevistaConsegna), "hh:nn")
''                                If DateDiff("n", varOraPrevistaConsegna, OraAttuale) > 0 Then
'                                If DateDiff("n", varOraPrevistaConsegna, OraAttuale) >= 50 Then
'                                    varOraPrevistaConsegna = Format(DateAdd("n", dblMinutiDelay, OraAttuale), "hh:nn")
'                                    Debug.Print "S" & " "; strTargaMotrice & " " & strIDViaggio & " " & strIDConsegna & " " & varOraPrevistaConsegna
'                                    With VVC
'
'                                        .DS.Fields.ResetFields
'                                        .DS.Fields("0").FindValue = strIDViaggio
'                                        .DS.Fields("1").FindValue = strIDConsegna
'
'                                        If Not .Find Then
'                                            objLogger.logError "Find Consegna."
'                                            GoTo ErrHandler
'                                        End If
'                                        .DS.Field("Orario previsto").NewValue = varOraPrevistaConsegna
'                                        .DS.Field("Orario previsto").Updated = True
'                                        .DS.Field("Flag stato 4").NewValue = "S"
'                                        .DS.Field("Flag stato 4").Updated = True
'
'                                        ' Salva la consegna
'                                        If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
'                                            objLogger.logError "Save Consegna."
'                                            GoTo ErrHandler
'                                        End If
'
'                                        .DS.DA.EndTrans etCommit
'                                    End With
'
'                                    ' Aggiorna il flag modificato (VVPFLST5) di tutti i prodotti del viaggio
'                                    If Not DS.DA.ExecuteCommand(SQL_UPDATE_VIAGGIO_MODIFICATO, strIDViaggio) Then
'                                        objLogger.logError "Errore in aggiornamento flag Modificato del viaggio."
'                                        GoTo ErrHandler
'                                    End If
'                                End If
'
'                                DS.DA.EndTrans etCommit
'                                j = k
'                            Next k
'                        End If
'                        lngPointer = j
'                    Next j
'                End If
'                i = lngPointer
'            End If
'        Next i
'    End If
'
'    AggiornaOrari_SMS = True
'
'    objLogger.logDebug "Fine AggiornaOrari_SMS senza errori."
'    GoTo EndProc
'
'ErrHandler:
'    DS.DA.EndTrans etRollback
'
''    Set VVP = Nothing
''    Set VVT = Nothing
''    Set VVC = Nothing
'
'    objLogger.logError "Fine AggiornaOrari_SMS con ERRORI."
'
'    AggiornaOrari_SMS = False
'
'EndProc:
'End Function

'-- DP 15/05/08
'-- Rifaccio la funzione viste le modifiche richieste
Public Function AggiornaOrari_SMS(Data As Variant, OraAttuale As Variant) As Boolean
    Dim intSogliaRitardo As Integer
    Dim intMinutiDelay As Integer
    Dim VVT As CViaggio
    Dim VVC As CViaggioConsegna
    Dim DataIeri As String
    Dim lngHowMany As Long
    Dim varBuffer As Variant
    Dim i As Long
    Dim lngBufferCount As Long
    Dim strTargaMotrice As String
    Dim strIDViaggio As String
    Dim strLastIDViaggio As String
    Dim strIDConsegna As String
    Dim varOraPrevistaCarico As Variant
    Dim varOraPrevistaConsegna As Variant
    
    On Error GoTo ErrHandler
    
    intSogliaRitardo = 50   ' Soglia di ritardo in minuti
    intMinutiDelay = 60     ' Minuti da aggiungere per carchi e/o consegne in ritardo di almeno [intSogliaRitardo]
    
    Set VVT = New CViaggio
    With VVT
        Set .DS.DA = DS.DA
        Set .DS.ErrH = DS.DA.ErrH
    End With
    
    Set VVC = New CViaggioConsegna
    With VVC
        Set .DS.DA = DS.DA
        Set .DS.ErrH = DS.DA.ErrH
    End With
    
    If IsDate(Data) Then
        DataIeri = Format(DateAdd("d", -1, CDate(Data)), "dd/mm/yyyy")
        Data = Format(Data, "dd/mm/yyyy")
    Else
        GoTo ErrHandler
    End If
    If IsDate(OraAttuale) Then
        OraAttuale = Format(OraAttuale, "hh:nn")
    Else
        GoTo ErrHandler
    End If
    
    lngHowMany = 0
    If Not DS.DA.QueryCommand(SQL_SMS_GET_LIST_AGG_2, varBuffer, lngHowMany, DataIeri, Data) Then
        objLogger.logError "Query SQL_SMS_GET_LIST_AGG_2."
        GoTo ErrHandler
    End If
    
    If Not IsEmpty(varBuffer) Then
        DS.DA.BeginTransaction
        
        lngBufferCount = UBound(varBuffer, 2)
        For i = 0 To lngBufferCount
            strTargaMotrice = CStr(varBuffer(0, i))
            strIDViaggio = CStr(varBuffer(1, i))
            strIDConsegna = CStr(varBuffer(2, i))
            varOraPrevistaCarico = Format(varBuffer(5, i), "hh:nn")
            varOraPrevistaConsegna = Format(varBuffer(6, i), "hh:nn")
            
            ' Se sono su un viaggio nuovo controllo se il carico è già stato effettuato
            If strLastIDViaggio <> strIDViaggio And (IsNull(varBuffer(3, i)) Or (varBuffer(3, i) <> "S")) Then
                ' Controllo se il carico è in ritardo rispetto all'orario passato come parametro
                If (DateDiff("n", varOraPrevistaCarico, OraAttuale) >= intSogliaRitardo) Then
                    ' Incremento varOraPrevistaCarico di intMinutiDelay
                    varOraPrevistaCarico = Format(DateAdd("n", intMinutiDelay, varOraPrevistaCarico), "hh:nn")
                        
                    ' Aggiorno l'ora di carico
                    With VVT
                        .DS.Fields.ResetFields
                        .DS.Fields("0").FindValue = strIDViaggio
                    
                        If Not .Find Then
                            objLogger.logError "Find Viaggio."
                            GoTo ErrHandler
                        End If
                        
                        .DS.Field("Ora inizio carico").NewValue = varOraPrevistaCarico
                        .DS.Field("Ora inizio carico").Updated = True
                        If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
                            objLogger.logError "Save Viaggio."
                            GoTo ErrHandler
                        End If
        
                        .DS.DA.EndTrans etCommit
                    End With
                End If
            End If
                        
            ' Controllo che la consegna non sia già iniziata (sulla carta questo controllo è inutile, ma non si sa mai...)
            If IsNull(varBuffer(4, i)) Or (varBuffer(4, i) <> "S") Then
                ' Controllo se la consegna è in ritardo rispetto all'orario passato come parametro
                If DateDiff("n", varOraPrevistaConsegna, OraAttuale) >= intSogliaRitardo Then
                    ' Incremento varOraPrevistaCarico di intMinutiDelay
                    varOraPrevistaConsegna = Format(DateAdd("n", intMinutiDelay, OraAttuale), "hh:nn")
                        
                    ' Aggiorno l'orario previsto di consegna
                    With VVC
                        .DS.Fields.ResetFields
                        .DS.Fields("0").FindValue = strIDViaggio
                        .DS.Fields("1").FindValue = strIDConsegna
                
                        If Not .Find Then
                            objLogger.logError "Find Consegna."
                            GoTo ErrHandler
                        End If
                        .DS.Field("Orario previsto").NewValue = varOraPrevistaConsegna
                        .DS.Field("Orario previsto").Updated = True
                        .DS.Field("Flag stato 4").NewValue = "S"
                        .DS.Field("Flag stato 4").Updated = True
                        
                        ' Salva la consegna
                        If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
                            objLogger.logError "Save Consegna."
                            GoTo ErrHandler
                        End If
                        
                        .DS.DA.EndTrans etCommit
                    End With
                    
                    ' Aggiorna il flag modificato (VVPFLST5) di tutti i prodotti del viaggio
                    If Not DS.DA.ExecuteCommand(SQL_UPDATE_VIAGGIO_MODIFICATO, strIDViaggio) Then
                        objLogger.logError "Errore in aggiornamento flag Modificato del viaggio."
                        GoTo ErrHandler
                    End If
                End If
            End If
            
            strLastIDViaggio = strIDViaggio
            DoEvents
        Next i
    
        DS.DA.EndTrans etCommit
    End If
    
    Set VVT = Nothing
    Set VVC = Nothing
    
    AggiornaOrari_SMS = True
    
    objLogger.logDebug "Fine AggiornaOrari_SMS senza errori."
    GoTo EndProc

ErrHandler:
    DS.DA.EndTrans etRollback
    
    Set VVT = Nothing
    Set VVC = Nothing
    
    objLogger.logError "Fine AggiornaOrari_SMS con ERRORI."
    AggiornaOrari_SMS = False

EndProc:

End Function


Public Function GetSelectionEvents(ViaggioId As String, Optional ActiveRecordsOnly As Boolean = False, _
                             Optional MaxRecords As Long = 0) As Boolean
    Dim strSQL As String

    strSQL = Replace(SQL_GET_SELECTION_EVENTS, "%sevidvvt%", ViaggioId)

    GetSelectionEvents = DS.GetSelection(strSQL, ActiveRecordsOnly, , MaxRecords)

End Function
Public Function UpdateKM(varIDViaggio As Variant, sngKmFineViaggio As Single)
    On Error GoTo ErrHandler
    
    Dim lngHowMany As Long
    Dim varBuffer As Variant
    Dim lngRowsCounter As Long
    Dim VVC As CViaggioConsegna
    Dim lngSequenzaConsegna As Long
    Dim lngKm As Long

    objLogger.logDebug "Inizio UpdateKM"

    ' Preleva i km di fine carico
    lngHowMany = 0
    If Not DS.DA.QueryCommand(SQL_GET_KM_EVENT, varBuffer, lngHowMany, varIDViaggio, "CF") Then
        objLogger.logError "Query SQL_GET_KM_EVENT."
        GoTo ErrHandler
    End If
    
    If lngHowMany = 0 Then
        objLogger.logError "Impossibile calcolare il km delle consegne. Manca l'evento fine carico."
        UpdateKM = True
        Exit Function
    End If

    lngKm = CLng(varBuffer(0, 0))
    
    ' Preleva i totali dei km
    lngHowMany = 0
    If Not DS.DA.QueryCommand(SQL_SELECT_KM_EVENTS, varBuffer, lngHowMany, varIDViaggio) Then
        objLogger.logError "Mancano le consegne."
        GoTo ErrHandler
    End If

    If lngHowMany > 0 Then

        Set VVC = New CViaggioConsegna
        With VVC
            Set .DS.DA = DS.DA
            Set .DS.ErrH = DS.DA.ErrH
        End With
        
        lngSequenzaConsegna = 10
        For lngRowsCounter = LBound(varBuffer, 2) To UBound(varBuffer, 2)

            objLogger.logDebug "ResOil Keys, VVTIDVVT = """ & varIDViaggio & _
                                """, VVCIDVVC = """ & varBuffer(0, lngRowsCounter) & """"

            ' Carica la consegna e aggiorna i km
            With VVC

                .DS.Fields.ResetFields
                .DS.Fields("0").FindValue = varIDViaggio
                .DS.Fields("1").FindValue = varBuffer(0, lngRowsCounter)

                If Not .Find Then
                    objLogger.logError "Consegna non trovata."
                    GoTo ErrHandler
                End If

                If .DS.HowMany > 0 Then
                    
                    ' Imposta i km
                    If lngRowsCounter = UBound(varBuffer, 2) Then
                        .DS.Fields("22").NewValue = sngKmFineViaggio - lngKm
                    Else
                        .DS.Fields("22").NewValue = CLng(varBuffer(2, lngRowsCounter)) - lngKm
                    End If
                    .DS.Fields("22").Updated = True
                    LogFieldValue .DS.Fields("22")

                    ' Imposta la sequenza del viaggio
                    .DS.Fields("2").NewValue = Format(lngSequenzaConsegna, "000")
                    .DS.Fields("2").Updated = True
                    LogFieldValue .DS.Fields("2")
    
                    ' Imposta i km a vuoto della consegna
                    If lngRowsCounter = UBound(varBuffer, 2) Then
                        .DS.Fields("45").NewValue = sngKmFineViaggio - CLng(varBuffer(2, lngRowsCounter))
                    Else
                        .DS.Fields("45").NewValue = 0
                    End If
                    .DS.Fields("45").Updated = True
                    LogFieldValue .DS.Fields("45")

                    ' Salva la consegna
                    If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
                        objLogger.logError "Save consegna."
                        GoTo ErrHandler
                    End If
    
                    .DS.DA.EndTrans etCommit
    
                End If
            
            End With

            lngKm = CLng(varBuffer(2, lngRowsCounter))
            lngSequenzaConsegna = lngSequenzaConsegna + 10
        Next lngRowsCounter
    End If
    
    objLogger.logDebug "Fine UpdateKM senza errori"
    UpdateKM = True
    Set VVC = Nothing
    Exit Function

ErrHandler:
    objLogger.logInfo "Fine UpdateKM con ERRORI"
    UpdateKM = False
    Set VVC = Nothing
End Function

Public Property Get Logger() As CLogger
    Set Logger = objLogger
End Property

Public Property Set Logger(objLoggerNew As CLogger)
    Set objLogger = objLoggerNew
End Property

Private Function IsEmptyVariantString(varValue As Variant) As Variant
    If IsNull(varValue) Or IsEmpty(varValue) Or Len(varValue) = 0 Then
        IsEmptyVariantString = True
    Else
        IsEmptyVariantString = False
    End If
End Function

'-- Restituisce true se le consegne con VVCNRSEQ inferiore hanno VVCFLST5 <> 'N' e da NULL
Public Function blnConsegnaInOrdine(varIDVVT As Variant, varNRSEQ As Variant) As Boolean
    Dim varBuffer As Variant
    
    On Error GoTo ErrHandler
    
    If DS.DA.QueryCommand(SQL_CONSEGNA_IN_ORDINE, varBuffer, 0, varIDVVT, varNRSEQ) Then
        If Not IsEmpty(varBuffer) Then
            GoTo ErrHandler
        End If
    Else
        GoTo ErrHandler
    End If
    
    blnConsegnaInOrdine = True
    Exit Function
    
ErrHandler:
    blnConsegnaInOrdine = False
End Function

' -- DP 10/07/2009
Function ScanEvents_OESIA(ByRef LastID As Variant) As Boolean
    Dim i As Long
    Dim varKBuffer As Variant
    Dim varBuffer As Variant
    Dim varKBuffer2 As Variant
    Dim strLastDate As String
    Dim lngHowMany As Long
    Dim varIDViaggio As Variant
    Dim varIDConsegna As Variant
    Dim varIDProdotto As Variant
    Dim VVP As CViaggioProdotto
    Dim VVT As CViaggio
    Dim VVC As CViaggioConsegna
    Dim ADZ As CDestinazione
    Dim dateLocalEventDateTime As Date
    Dim sngKm As Single
    Dim strSQLStatement As String
    Dim blnUsaDatiGPS As Boolean
'    Dim blnAttesaPedaggio As Boolean

    Dim blnViaggioNonTerminato As Boolean
    
    Dim verBuffer As Variant
    Dim varViaggiGiorno As Variant

    objLogger.logDebug "Inizio ScanEvents_OESIA."
    
 
       
     
    ' Scan KNOSOS table events
    lngHowMany = 0
    If Not DS.DA.QueryCommand(SQL_OESIA_GET_EVENTS, varKBuffer, lngHowMany, LastID) Then
        objLogger.logError "Query SQL_OESIA_GET_EVENTS."
        GoTo ErrHandler
    End If

    Set VVT = New CViaggio
    With VVT
        Set .DS.DA = DS.DA
        Set .DS.ErrH = DS.DA.ErrH
    End With
    
    Set VVP = New CViaggioProdotto
    With VVP
        Set .DS.DA = DS.DA
        Set .DS.ErrH = DS.DA.ErrH
    End With
    
    Set VVC = New CViaggioConsegna
    With VVC
        Set .DS.DA = DS.DA
        Set .DS.ErrH = DS.DA.ErrH
    End With
    
    Set ADZ = New CDestinazione
    With ADZ
        Set .DS.DA = DS.DA
        Set .DS.ErrH = DS.DA.ErrH
    End With
    
    objLogger.logInfo "Trovati " & lngHowMany & " eventi."
    
    If lngHowMany > 0 Then
        For i = LBound(varKBuffer, 2) To UBound(varKBuffer, 2)
            ' Inizia transazione
            DS.DA.BeginTransaction

'            blnAttesaPedaggio = False

            varIDViaggio = ""
            varIDConsegna = ""
            varIDProdotto = ""

            objLogger.logInfo "Evento n. " & i + 1
            objLogger.logInfo "-------------------------------------------------"
            objLogger.logInfo "CodigoVehiculo [" & varKBuffer(0, i) & "]"
            objLogger.logInfo "CodigoViaje    [" & varKBuffer(1, i) & "]"
            objLogger.logInfo "CodigoServicio [" & varKBuffer(2, i) & "]"
            objLogger.logInfo "FechaEvento    [" & varKBuffer(3, i) & "]"
            objLogger.logInfo "CodigoEvento   [" & varKBuffer(4, i) & "]"
            objLogger.logInfo "Metros         [" & varKBuffer(5, i) & "]"
            objLogger.logInfo "Coordinata X   [" & varKBuffer(9, i) & "]"
            objLogger.logInfo "Coordinata Y   [" & varKBuffer(8, i) & "]"
            objLogger.logInfo "-------------------------------------------------"
            
            ' E' un evento relativo al viaggio o carico o scarico?
            Select Case UCase(varKBuffer(4, i))
                Case "ON"   ' Inicio trabajo
                    GoTo NextIteration
                
                Case "LLE", "SAL", "OFF"   ' Llegada a la factoria - Salida de factoria
                    lngHowMany = 0
                    If Not DS.DA.QueryCommand(SQL_KNOSOS_GET_KEY_VIAGGIO_RESOIL, varBuffer, lngHowMany, CStr(varKBuffer(2, i))) Then
                        objLogger.logError "Query SQL_KNOSOS_GET_KEY_VIAGGIO_RESOIL."
                        GoTo ErrHandler
                    End If
                    
                    If lngHowMany > 0 Then
                        varIDViaggio = varBuffer(0, 0)
                        varIDConsegna = ""
                        varIDProdotto = ""
                    Else
                        objLogger.logWarning "Viaggio non trovato in ResOil."
                        GoTo NextIteration
                    End If
                
                Case "INI", "FIN"   ' Inicio descarga - Final descarga
                    lngHowMany = 0
                    If Not DS.DA.QueryCommand(SQL_KNOSOS_GET_KEY_CONSEGNA_PRODOTTO_RESOIL, varBuffer, lngHowMany, CStr(varKBuffer(2, i)), CVar(varKBuffer(1, i))) Then
                        objLogger.logError "Query SQL_KNOSOS_GET_KEY_CONSEGNA_PRODOTTO_RESOIL."
                        GoTo ErrHandler
                    End If
                    
                    If lngHowMany > 0 Then
                        varIDViaggio = varBuffer(0, 0)
                        varIDConsegna = varBuffer(1, 0)
                        varIDProdotto = varBuffer(2, 0)
                    Else
                        objLogger.logWarning "Consegna prodotto non trovato in ResOil."
                        GoTo NextIteration
                    End If
                    
'                Case "OFF"  ' Fin trabajo
'                    ' Trovo l'ultimo viaggio caricato del giorno
'                    varViaggiGiorno = Empty
'                    varViaggiGiorno = varGetViaggiCaricatiDelGiorno(VVT.DS.Fields("14").Value, Format(varKBuffer(3, i), "dd/mm/yyyy"))
'                    If Not IsEmpty(varViaggiGiorno) Then
'                        varIDViaggio = varViaggiGiorno(0, UBound(varViaggiGiorno, 2))
'                        varIDConsegna = ""
'                        varIDProdotto = ""
'                    Else
'                        objLogger.logWarning "Non è stato trovato nessun viaggio caricato per questa UdT in questo giorno, evento OFF ignorato"
'                        GoTo NextIteration
'                    End If
                    
                Case Else
                    objLogger.logError "Evento " & UCase(varKBuffer(4, i)) & " sconosciuto."
                    GoTo ErrHandler
                    
            End Select
            
            objLogger.logInfo "ResOil Keys VVTIDVVT = """ & varIDViaggio & _
                                """, VVCIDVVC = """ & varIDConsegna & """, VVPIDVVP = """ & varIDProdotto & """"
            
            ' Log record event
            DS.DA.LogExecTime "INFO: KNOSOS Evento", -1
            DS.DA.LogExecTime "INFO: -----------------------------------------------------------------", -1
            DS.DA.LogExecTime "INFO: - CodigoVehiculo [" & varKBuffer(0, i) & "]", -1
            DS.DA.LogExecTime "INFO: - CodigoViaje    [" & varKBuffer(1, i) & "]", -1
            DS.DA.LogExecTime "INFO: - CodigoServicio [" & varKBuffer(2, i) & "]", -1
            DS.DA.LogExecTime "INFO: - FechaEvento    [" & varKBuffer(3, i) & "]", -1
            DS.DA.LogExecTime "INFO: - CodigoEvento   [" & varKBuffer(4, i) & "]", -1
            DS.DA.LogExecTime "INFO: - Metros         [" & varKBuffer(5, i) & "]", -1
            DS.DA.LogExecTime "INFO: - Coordinata X   [" & varKBuffer(9, i) & "]", -1
            DS.DA.LogExecTime "INFO: - Coordinata Y   [" & varKBuffer(8, i) & "]", -1

            ' Se non trovo il viaggio o prodotto, skip
            If varIDViaggio <> "" Then
                dateLocalEventDateTime = StringKNOSOSDate_To_Date(CStr(varKBuffer(6, i)))
                If lngHowMany > 0 Then
                    DS.Fields.ResetFields
            
                    ' Creazione dell'evento
                    DS.Fields("0").NewValue = DS.DA.GetCounter("anaganu0", "anuprogr", "anuidanu", "901", 0)
                    DS.Fields("0").Updated = True
            
                    DS.Fields("1").NewValue = varIDViaggio
                    DS.Fields("1").Updated = True
        
                    DS.Fields("2").NewValue = varIDConsegna
                    DS.Fields("2").Updated = True
        
                    DS.Fields("3").NewValue = varIDProdotto
                    DS.Fields("3").Updated = True
        
                    DS.Fields("5").NewValue = varKBuffer(6, i)
                    DS.Fields("5").Updated = True
        
'                    DS.Fields("6").NewValue = varKBuffer(7, i)
'                    DS.Fields("6").Updated = True
'
'                    DS.Fields("7").NewValue = varKBuffer(8, i)
'                    DS.Fields("7").Updated = True
        
                    DS.Fields("8").NewValue = CDbl(varKBuffer(5, i)) / 1000
                    DS.Fields("8").Updated = True
                    
                    DS.Fields("9").NewValue = "K"
                    DS.Fields("9").Updated = True
        
                    ' Carica il viaggio
                    With VVT
                        .DS.Fields.ResetFields
                        .DS.Fields("0").FindValue = varIDViaggio
                    
                        If Not .Find Then
                            objLogger.logError "Find Viaggio."
                            GoTo ErrHandler
                        End If
                        
                        If IsEmptyVariantString(.DS.Field("Flag Usa Dati GPS").Value) Or _
                            .DS.Field("Flag Usa Dati GPS").Value = "S" Then
                            blnUsaDatiGPS = True
                        Else
                            blnUsaDatiGPS = False
                        End If
                    
                    End With
                    
                    Select Case UCase(varKBuffer(4, i))
'                        ' -------------------------------------------------------
'                        ' Evento - Inicio del viaggio (Inizio del viaggio)
'                        ' -------------------------------------------------------
                        Case "ON "
'                            DS.Fields("4").NewValue = "VI"
'                            DS.Fields("4").Updated = True
'
'                            objLogger.logInfo "Tipo evento: Inizio viaggio"
'
'                            ' Aggiorna lo stato del viaggio (VVTFLSTA)
'                            VVT.DS.Fields("48").NewValue = VIAGGIO_STATO_INIZIOVIAGGIO
'                            VVT.DS.Fields("48").Updated = True
'                            LogFieldValue VVT.DS.Fields("48")
'
'                            If blnUsaDatiGPS Then
'                                ' Aggiorna il flag modificato (VVPFLST5) di tutti i prodotti del viaggio
'                                If Not DS.DA.ExecuteCommand(SQL_UPDATE_VIAGGIO_MODIFICATO, varIDViaggio) Then
'                                    objLogger.logError "Errore nell'aggiornamento del flag Modificato."
'                                    GoTo ErrHandler
'                                End If
'                            End If

                        ' -------------------------------------------------------
                        ' Evento - Llegada en factora (Inizio carico)
                        ' -------------------------------------------------------
                        Case "LLE"
                            DS.Fields("4").NewValue = "CI"
                            DS.Fields("4").Updated = True
        
                            objLogger.logInfo "Tipo evento: Inizio carico"
                            
                            If CDbl(varKBuffer(5, i)) > 0 Then
                                ' Controllo di non avere altri eventi LLE in precedenza per questo stesso viaggio
                                varBuffer = Empty
                                If DS.DA.QueryCommand(SQL_GET_EVENTI_PRECEDENTI, varBuffer, 0, CVar(varKBuffer(4, i)), CVar(varKBuffer(2, i)), CVar(varKBuffer(7, i))) Then
                                    If IsEmpty(verBuffer) Then
                                        ' Se è il primo viaggio gli aggiorno i "Km a vuoto", altrimenti aggiorno i "Km di rientro reali" del viaggio precedente
                                        varViaggiGiorno = Empty
                                        varViaggiGiorno = varGetViaggiCaricatiDelGiorno(VVT.DS.Fields("14").Value, VVT.DS.Fields("6").Value)
                                        If IsEmpty(varViaggiGiorno) Then
                                            ' Non ho altri viaggi caricati in questo giorno per questa UdT, aggiorno i "Km a vuoto" per il viaggio in corso
                                            VVT.DS.Fields("49").NewValue = CDbl(varKBuffer(5, i)) / 1000
                                            VVT.DS.Fields("49").Updated = True
                                            objLogger.logInfo "Rilevato primo viaggio, aggiorno Km a vuoto"
                                            LogFieldValue VVT.DS.Fields("49")
                                            
                '                            ' Aggiorna lo stato del viaggio (VVTFLSTA)
                '                            VVT.DS.Fields("48").NewValue = VIAGGIO_STATO_CARICO
                '                            VVT.DS.Fields("48").Updated = True
                '                            LogFieldValue VVT.DS.Fields("48")
                '
                '                            ' Aggiorna flag carico completato
                '                            VVT.DS.Fields("52").NewValue = "S"
                '                            VVT.DS.Fields("52").Updated = True
                '                            LogFieldValue VVT.DS.Fields("52")
                '
                '                            If blnUsaDatiGPS Then
                '                                ' Aggiorna l'ora inizio carico sul viaggio (VVTHRINC)
                '                                VVT.DS.Fields("5").NewValue = TruncDate(dateLocalEventDateTime)
                '                                VVT.DS.Fields("5").Updated = True
                '                                LogFieldValue VVT.DS.Fields("5")
                '
                '                                ' Aggiorna il flag modificato (VVPFLST5) di tutti i prodotti del viaggio
                '                                If Not DS.DA.ExecuteCommand(SQL_UPDATE_VIAGGIO_MODIFICATO, varIDViaggio) Then
                '                                    objLogger.logError "Errore nell'aggiornamento del flag Modificato."
                '                                    GoTo ErrHandler
                '                                End If
                '                            End If
                                        Else
                                            ' Ho altri viaggi caricati in questo giorno per questa UdT, aggiorno i "Km di rientro reali" del viaggio precedente
                                            ' Carica il viaggio
                                            objLogger.logInfo "Rilevati altri viaggi caricati nel giorno, aggiorno Km di rientro del viaggio precedente"
                                            
                                            VVT.DS.Fields("49").NewValue = 0
                                            VVT.DS.Fields("49").Updated = True
                                            objLogger.logInfo "Azzerati i Km a vuoto del viaggio"
                                            LogFieldValue VVT.DS.Fields("49")
                                            
                                            With VVT
                                                .DS.Fields.ResetFields
                                                .DS.Fields("0").FindValue = varViaggiGiorno(0, UBound(varViaggiGiorno, 2))
                                            
                                                If Not .Find Then
                                                    objLogger.logError "Find Viaggio."
                                                    GoTo ErrHandler
                                                End If
                                                
                                                If IsEmptyVariantString(.DS.Field("Flag Usa Dati GPS").Value) Or _
                                                    .DS.Field("Flag Usa Dati GPS").Value = "S" Then
                                                    blnUsaDatiGPS = True
                                                Else
                                                    blnUsaDatiGPS = False
                                                End If
                                            
                                                ' Aggiorno i Km di rientro reali per il viaggio
                                                VVT.DS.Fields("57").NewValue = CDbl(varKBuffer(5, i)) / 1000
                                                VVT.DS.Fields("57").Updated = True
                                                LogFieldValue VVT.DS.Fields("57")
                                                objLogger.logInfo "Aggiornati i Km di rientro reali del viaggio " & varViaggiGiorno(0, UBound(varViaggiGiorno, 2))
                                            End With
                                        End If
                                    Else
                                        objLogger.logWarning "Evento LLE ignorato in quanto ne è già stato ricevuto un altro in precedenza per lo stesso viaggio"
                                    End If
                                Else
                                    objLogger.logError "Errore durante il controllo sull'evento LLE"
                                End If
                            Else
                                objLogger.logWarning "Evento LLE ignorato in quanto giunto con i Km a 0"
                            End If
    
                        ' -------------------------------------------------------
                        ' Evento - Salida de factora (Fine del carico)
                        ' -------------------------------------------------------
                        Case "SAL"
                            DS.Fields("4").NewValue = "CF"
                            DS.Fields("4").Updated = True
        
                            objLogger.logInfo "Tipo evento: Fine carico"
                            
                            ' Controllo di non avere altri eventi SAL in precedenza per questo stesso viaggio
                            varBuffer = Empty
                            If DS.DA.QueryCommand(SQL_GET_EVENTI_PRECEDENTI, varBuffer, 0, CVar(varKBuffer(4, i)), CVar(varKBuffer(2, i)), CVar(varKBuffer(7, i))) Then
                                If IsEmpty(verBuffer) Then
        '                            ' Aggiorna lo stato del viaggio (VVTFLSTA)
        '                            VVT.DS.Fields("48").NewValue = VIAGGIO_STATO_INCORSO
        '                            VVT.DS.Fields("48").Updated = True
        '                            LogFieldValue VVT.DS.Fields("48")
        
                                    ' Aggiorna il tempo di carico sul viaggio (VVTTMPCA)
'                                    If blnUsaDatiGPS Then
                                        VVT.DS.Fields("30").NewValue = DateDiff("n", TruncDate(VVT.DS.Fields("5").Value), TruncDate(dateLocalEventDateTime))
                                        VVT.DS.Fields("30").Updated = True
                                        LogFieldValue VVT.DS.Fields("30")
        
        '                                ' Aggiorna l'ora inizio viaggio sul viaggio (VVTHRINV)
        '                                VVT.DS.Fields("7").NewValue = TruncDate(dateLocalEventDateTime)
        '                                VVT.DS.Fields("7").Updated = True
        '                                LogFieldValue VVT.DS.Fields("7")
        '
        '                                ' Aggiorna flag carico completato
        '                                VVT.DS.Fields("39").NewValue = "S"
        '                                VVT.DS.Fields("39").Updated = True
        '                                LogFieldValue VVT.DS.Fields("39")
        '
        '                                ' Aggiorna il flag modificato (VVPFLST5) di tutti i prodotti della consegna
        '                                If Not DS.DA.ExecuteCommand(SQL_UPDATE_VIAGGIO_MODIFICATO, varIDViaggio) Then
        '                                    objLogger.logError "Errore in aggiornamento flag Modificato del viaggio."
        '                                    GoTo ErrHandler
        '                                End If
'                                    End If
                                Else
                                    objLogger.logWarning "Evento SAL ignorato in quanto ne è già stato ricevuto un altro in precedenza per lo stesso viaggio"
                                End If
                            Else
                                objLogger.logError "Errore durante il controllo sull'evento SAL"
                            End If

                        ' -------------------------------------------------------
                        ' Evento - Inicio descarga (Inizio scarico prodotto)
                        ' -------------------------------------------------------
                        Case "INI"
                            DS.Fields("4").NewValue = "SI"
                            DS.Fields("4").Updated = True
        
                            objLogger.logInfo "Tipo evento: Inizio consegna prodotto"
                            
                            ' Carica la consegna prodotto e aggiorna i dati
'                            With VVP
'                                .DS.Fields.ResetFields
'                                .DS.Fields("0").FindValue = varIDViaggio
'                                .DS.Fields("1").FindValue = varIDConsegna
'                                .DS.Fields("2").FindValue = varIDProdotto
'
'                                If Not .Find Then
'                                    objLogger.logError "Find Consegna Prodotto."
'                                    GoTo ErrHandler
'                                End If
'
'                                ' Aggiorna lo stato dello scarico (VVPFLSTA)
'                                .DS.Fields("37").NewValue = "1"
'                                .DS.Fields("37").Updated = True
'                                LogFieldValue .DS.Fields("37")
'
'                                ' Aggiorna il flag Modificato (VVPFLST5)
'                                .DS.Field("Stato Riga 5").NewValue = "S"
'                                .DS.Field("Stato Riga 5").Updated = True
'                                LogFieldValue .DS.Field("Stato Riga 5")
'
'                                ' Aggiorna l'ora inizio scarico (VVPHRINC)
'                                If blnUsaDatiGPS Then
'                                    .DS.Fields("38").NewValue = TruncDate(dateLocalEventDateTime)
'                                    .DS.Fields("38").Updated = True
'                                    LogFieldValue .DS.Fields("38")
'                                End If
'
'                                ' Salva
'                                If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
'                                    objLogger.logError "Save Consegna Prodotto."
'                                    GoTo ErrHandler
'                                End If
'
'                                .DS.DA.EndTrans etCommit
'                            End With
                            
                            ' Carica la consegna e aggiorna i dati
                            With VVC
                                .DS.Fields.ResetFields
                                .DS.Fields("0").FindValue = varIDViaggio
                                .DS.Fields("1").FindValue = varIDConsegna
                        
                                If Not .Find Then
                                    objLogger.logError "Find Consegna."
                                    GoTo ErrHandler
                                End If
                                
'                                ' E' il primo scarico della consegna
'                                If .DS.Fields("44").Value = "0" Or IsNull(.DS.Fields("44").Value) Or .DS.Fields("44").Value = "" Then
'                                    If blnUsaDatiGPS Then
'                                        ' Aggiorna l'ora di arrivo
'                                        .DS.Fields("24").NewValue = TruncDate(dateLocalEventDateTime)
'                                        .DS.Fields("24").Updated = True
'                                        LogFieldValue .DS.Fields("24")
'
'                                        ' Aggiorna il flag consegna iniziata
'                                        .DS.Fields("36").NewValue = "S"
'                                        .DS.Fields("36").Updated = True
'                                        LogFieldValue .DS.Fields("36")
'                                    End If
'
'                                    ' Aggiorna lo stato 'In corso'
'                                    .DS.Fields("44").NewValue = "1"
'                                    .DS.Fields("44").Updated = True
'                                    LogFieldValue .DS.Fields("44")
'                                End If

                                ' Imposta i km parziali
                                If Not (IsNull(varKBuffer(5, i)) Or IsEmpty(varKBuffer(5, i))) Then
                                    .DS.Fields("22").NewValue = CDbl(varKBuffer(5, i)) / 1000
                                    .DS.Fields("22").Updated = True
                                End If
                                
                                ' Salva la consegna
                                If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
                                    objLogger.logError "Save Consegna."
                                    GoTo ErrHandler
                                End If
                                
                                .DS.DA.EndTrans etCommit
                            End With
                            
                            ' Carica la destinazione e aggiorna le coordinate
                            With ADZ
                                .DS.Fields("0").FindValue = VVC.DS.Fields("12").Value
                            
                                If Not .Find Then
                                    objLogger.logError "Find Destinazione."
                                    GoTo ErrHandler
                                End If
                                
                                .DS.Fields("31").NewValue = CVar(varKBuffer(9, i))
                                .DS.Fields("31").Updated = True
                                .DS.Fields("32").NewValue = CVar(varKBuffer(8, i))
                                .DS.Fields("32").Updated = True
                            
                                If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
                                    objLogger.logError "Save Destinazione."
                                    GoTo ErrHandler
                                End If
                                
                                .DS.DA.EndTrans etCommit
                            End With
                            
                            
'                            ' Se ho ricevuto tutti gli scarichi per il viaggio corrente aggiorno anche i "Km definitivi" del viaggio
'                            varBuffer = Empty
'                            blnViaggioNonTerminato = False
'                            If DS.DA.QueryCommand(SQL_GET_SCARICHI_NON_CONCLUSI, varBuffer, 0, strIDVVT) Then
'                                If Not IsEmpty(varBuffer) Then
'                                    For lngHowMany = 0 To UBound(varBuffer, 2)
'                                        If IsNull(varBuffer(2, lngHowMany)) Then
'                                            blnViaggioNonTerminato = True
'                                            Exit For
'                                        End If
'                                    Next
'
'                                    If Not blnViaggioNonTerminato Then
'                                        ' Sono stati effettuati tutti gli scarichi
'                                        varBuffer = Empty
'                                        If DS.DA.QueryCommand(SQL_GET_KM_VIAGGIO, varBuffer, 0, strIDVVT) Then
'                                            VVT.DS.Fields("27").NewValue = CDbl(varBuffer(0, 0)) / 1000
'                                            VVT.DS.Fields("27").Updated = True
'                                            LogFieldValue VVT.DS.Fields("27")
'                                        Else
'                                            objLogger.logError "Errori in SQL_GET_KM_VIAGGIO"
'                                            GoTo ErrHandler
'                                        End If
'                                    End If
'                                Else
'                                    objLogger.logError "Errori in SQL_GET_SCARICHI_NON_CONCLUSI: zero risultati"
'                                    GoTo ErrHandler
'                                End If
'                            Else
'                                objLogger.logError "Errori in SQL_GET_SCARICHI_NON_CONCLUSI"
'                                GoTo ErrHandler
'                            End If
    
                        ' -------------------------------------------------------
                        ' Evento - Fin descarga (Fine scarico prodotto)
                        ' -------------------------------------------------------
                        Case "FIN"
                            DS.Fields("4").NewValue = "SF"
                            DS.Fields("4").Updated = True
                            
                            objLogger.logInfo "Tipo evento: Fine consegna prodotto"
                            
'                            ' Carica la consegna prodotto e aggiorna i dati
'                            With VVP
'                                .DS.Fields.ResetFields
'                                .DS.Fields("0").FindValue = varIDViaggio
'                                .DS.Fields("1").FindValue = varIDConsegna
'                                .DS.Fields("2").FindValue = varIDProdotto
'
'                                If Not .Find Then
'                                    objLogger.logError "Find Consegna Prodotto."
'                                    GoTo ErrHandler
'                                End If
'
'                                ' Aggiorna lo stato dello scarico (VVPFLSTA)
'                                .DS.Fields("37").NewValue = "2"
'                                .DS.Fields("37").Updated = True
'                                LogFieldValue .DS.Fields("37")
'
'                                If blnUsaDatiGPS Then
'                                    ' Imposta i litri scaricati
''                                    If IsNull(varKBuffer(8, i)) Or IsEmpty(varKBuffer(8, i)) Then
''                                        .DS.Fields("11").NewValue = 0
''                                    Else
''                                        .DS.Fields("11").NewValue = CSng(varKBuffer(8, i))
''                                    End If
''                                    .DS.Fields("11").Updated = True
''                                    LogFieldValue .DS.Fields("11")
'
'                                    ' Imposta i litri caricati
''                                    .DS.Fields("7").NewValue = .DS.Fields("11").NewValue
''                                    .DS.Fields("7").Updated = True
''                                    LogFieldValue .DS.Fields("7")
'
'                                    ' Aggiorna l'ora fine scarico (VVPHRFIC)
'                                    .DS.Fields("39").NewValue = TruncDate(dateLocalEventDateTime)
'                                    .DS.Fields("39").Updated = True
'                                    LogFieldValue .DS.Fields("39")
'
'                                    ' Aggiorna il flag Modificato (VVPFLST5)
'                                    .DS.Field("Stato Riga 5").NewValue = "S"
'                                    .DS.Field("Stato Riga 5").Updated = True
'                                    LogFieldValue .DS.Field("Stato Riga 5")
'                                End If
'
'                                ' Salva
'                                If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
'                                    objLogger.logError "Save Consegna Prodotto."
'                                    GoTo ErrHandler
'                                End If
'
'                                .DS.DA.EndTrans etCommit
'                            End With
'
'                            ' Carica la consegna e aggiorna i dati
'                            With VVC
'                                .DS.Fields.ResetFields
'                                .DS.Fields("0").FindValue = varIDViaggio
'                                .DS.Fields("1").FindValue = varIDConsegna
'
'                                If Not .Find Then
'                                    objLogger.logError "Find Consegna."
'                                    GoTo ErrHandler
'                                End If
'
'                                ' Se e' l'ultimo scarico della consegna, aggiorna il tempo scarico e lo stato
'                                lngHowMany = 0
'                                If Not DS.DA.QueryCommand(SQL_KNOSOS_GET_COUNT_SCARICHI_NONCONCLUSI, varBuffer, lngHowMany, _
'                                    varIDViaggio, varIDConsegna) Then
'                                    objLogger.logError "Query SQL_KNOSOS_GET_COUNT_SCARICHI_NONCONCLUSI."
'                                    GoTo ErrHandler
'                                End If
'
'                                If CLng(varBuffer(0, 0)) = 0 Then
'                                    If blnUsaDatiGPS Then
'                                        ' Aggiorna il tempo di scarico
'                                        .DS.Fields("25").NewValue = DateDiff("n", TruncDate(.DS.Fields("24").Value), TruncDate(dateLocalEventDateTime))
'                                        .DS.Fields("25").Updated = True
'                                        LogFieldValue .DS.Fields("25")
'
'                                        ' Aggiorna il flag consegna ultimata
'                                        .DS.Fields("47").NewValue = "S"
'                                        .DS.Fields("47").Updated = True
'                                        LogFieldValue .DS.Fields("47")
'                                    End If
'
'                                    ' Aggiorna lo stato 'Completo'
'                                    .DS.Fields("44").NewValue = "2"
'                                    .DS.Fields("44").Updated = True
'                                    LogFieldValue .DS.Fields("44")
'                                End If
'
'                                ' Imposta i km parziali
'                                If Not (IsNull(varKBuffer(9, i)) Or IsEmpty(varKBuffer(9, i))) Then
'                                    .DS.Fields("22").NewValue = GetKmConsegna(varIDViaggio, varIDConsegna) + CSng(varKBuffer(9, i))
'                                    .DS.Fields("22").Updated = True
'                                End If
'
'                                ' Salva la consegna
'                                If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
'                                    objLogger.logError "Save Consegna."
'                                    GoTo ErrHandler
'                                End If
'
'                                .DS.DA.EndTrans etCommit
'                            End With
                        
'                        ' -------------------------------------------------------
'                        ' Evento - Rechazo de descarga (Dirottamento scarico)
'                        ' -------------------------------------------------------
'                        Case "REC"
'                            DS.Fields("4").NewValue = "SA"
'                            DS.Fields("4").Updated = True
'
'                            objLogger.logInfo "Tipo evento: Dirottamento o Rientro"
'
'                            ' Carica la consegna
'                            With VVC
'
'                                .DS.Fields.ResetFields
'                                .DS.Fields("0").FindValue = varIDViaggio
'                                .DS.Fields("1").FindValue = varIDConsegna
'
'                                If Not .Find Then
'                                    objLogger.logError "Find Consegna."
'                                    GoTo ErrHandler
'                                End If
'
'                            End With
'
'                            ' Carica la consegna prodotto e aggiorna i dati
'                            With VVP
'                                .DS.Fields.ResetFields
'                                .DS.Fields("0").FindValue = varIDViaggio
'                                .DS.Fields("1").FindValue = varIDConsegna
'                                .DS.Fields("2").FindValue = varIDProdotto
'
'                                If Not .Find Then
'                                    objLogger.logError "Find Consegna Prodotto."
'                                    GoTo ErrHandler
'                                End If
'
'                                ' Aggiorna lo stato dello scarico (VVPFLSTA)
'                                .DS.Fields("37").NewValue = "2"
'                                .DS.Fields("37").Updated = True
'                                LogFieldValue .DS.Fields("37")
'
'                                ' Imposta i litri caricati e scaricati
'                                Select Case VVC.DS.Fields("33").Value
'
'                                    ' Dirottamento
'                                    Case "D"
'                                        objLogger.logInfo "Tipo evento: Dirottamento"
'                                        If IsNull(varKBuffer(8, i)) Or IsEmpty(varKBuffer(8, i)) Then
'                                            .DS.Fields("11").NewValue = 0
'                                        Else
'                                            .DS.Fields("11").NewValue = CSng(varKBuffer(8, i))
'                                        End If
'                                        .DS.Fields("11").Updated = True
'                                        LogFieldValue .DS.Fields("11")
'
'                                        ' Imposta i litri caricati
'                                        If blnUsaDatiGPS Then
'                                            .DS.Fields("7").NewValue = 0
'                                            .DS.Fields("7").Updated = True
'                                            LogFieldValue .DS.Fields("7")
'                                        End If
'
'                                    ' Rientro
'                                    Case "R"
'                                        objLogger.logInfo "Tipo evento: Rientro"
'
'                                        If blnUsaDatiGPS Then
'                                            ' Imposta i litri scaricati
'                                            If IsNull(varKBuffer(7, i)) Or IsEmpty(varKBuffer(7, i)) Then
'                                                .DS.Fields("11").NewValue = 0
'                                            Else
'                                                .DS.Fields("11").NewValue = CSng(varKBuffer(7, i))
'                                            End If
'                                            .DS.Fields("11").Updated = True
'                                            LogFieldValue .DS.Fields("11")
'
'                                            ' Imposta i litri caricati
'                                            .DS.Fields("7").NewValue = .DS.Fields("11").NewValue
'                                            .DS.Fields("7").Updated = True
'                                            LogFieldValue .DS.Fields("7")
'                                        End If
'
'                                End Select
'
'                                If blnUsaDatiGPS Then
'                                    ' Aggiorna l'ora fine scarico (VVPHRFIC)
'                                    .DS.Fields("39").NewValue = TruncDate(dateLocalEventDateTime)
'                                    .DS.Fields("39").Updated = True
'                                    LogFieldValue .DS.Fields("39")
'
'                                    ' Aggiorna il flag Modificato (VVPFLST5)
'                                    .DS.Field("Stato Riga 5").NewValue = "S"
'                                    .DS.Field("Stato Riga 5").Updated = True
'                                    LogFieldValue .DS.Field("Stato Riga 5")
'                                End If
'
'                                ' Salva
'                                If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
'                                    objLogger.logError "Save Consegna Prodotto."
'                                    GoTo ErrHandler
'                                End If
'
'                                .DS.DA.EndTrans etCommit
'                            End With
'
'                            ' Carica la consegna e aggiorna i dati
'                            With VVC
'
'                                ' Se e' l'ultimo scarico della consegna, aggiorna il tempo scarico e lo stato
'                                lngHowMany = 0
'                                If Not DS.DA.QueryCommand(SQL_KNOSOS_GET_COUNT_SCARICHI_NONCONCLUSI, varBuffer, lngHowMany, _
'                                    varIDViaggio, varIDConsegna) Then
'                                    objLogger.logError "Query SQL_KNOSOS_GET_COUNT_SCARICHI_NONCONCLUSI."
'                                    GoTo ErrHandler
'                                End If
'
'                                If CLng(varBuffer(0, 0)) = 0 Then
'
'                                    If blnUsaDatiGPS Then
'                                        ' Aggiorna il tempo di scarico
'                                        .DS.Fields("25").NewValue = DateDiff("n", TruncDate(.DS.Fields("24").Value), TruncDate(dateLocalEventDateTime))
'                                        .DS.Fields("25").Updated = True
'                                        LogFieldValue .DS.Fields("25")
'
'                                        ' Aggiorna il flag consegna ultimata
'                                        .DS.Fields("36").NewValue = "S"
'                                        .DS.Fields("36").Updated = True
'                                        LogFieldValue .DS.Fields("36")
'                                    End If
'
'                                    ' Aggiorna lo stato 'Completo'
'                                    .DS.Fields("44").NewValue = "2"
'                                    .DS.Fields("44").Updated = True
'                                    LogFieldValue .DS.Fields("44")
'                                End If
'
''                                ' Imposta i km parziali
''                                If Not (IsNull(varKBuffer(9, i)) Or IsEmpty(varKBuffer(9, i))) Then
''                                    .DS.Fields("22").NewValue = GetKmConsegna(varIDViaggio, varIDConsegna) + CSng(varKBuffer(9, i))
''                                    .DS.Fields("22").Updated = True
''                                End If
'
'                                ' Salva la consegna
'                                If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
'                                    objLogger.logError "Save Consegna."
'                                    GoTo ErrHandler
'                                End If
'
'                                .DS.DA.EndTrans etCommit
'                            End With
'
'                        ' -------------------------------------------------------
'                        ' Evento - Devolucin de la descarga (Rientro scarico)
'                        ' -------------------------------------------------------
'                        Case "DEV"
'                            DS.Fields("4").NewValue = "SD"
'                            DS.Fields("4").Updated = True
'
'                            objLogger.logInfo "Tipo evento: Rientro"
'
'                            ' Carica la consegna prodotto e aggiorna i dati
'                            With VVP
'                                .DS.Fields.ResetFields
'                                .DS.Fields("0").FindValue = varIDViaggio
'                                .DS.Fields("1").FindValue = varIDConsegna
'                                .DS.Fields("2").FindValue = varIDProdotto
'
'                                If Not .Find Then
'                                    objLogger.logError "Find Consegna Prodotto."
'                                    GoTo ErrHandler
'                                End If
'
'                                ' Aggiorna lo stato dello scarico (VVPFLSTA)
'                                .DS.Fields("37").NewValue = "2"
'                                .DS.Fields("37").Updated = True
'                                LogFieldValue .DS.Fields("37")
'
'                                If blnUsaDatiGPS Then
'                                    ' Imposta i litri scaricati
'                                    If IsNull(varKBuffer(7, i)) Or IsEmpty(varKBuffer(7, i)) Then
'                                        .DS.Fields("11").NewValue = 0
'                                    Else
'                                        .DS.Fields("11").NewValue = CSng(varKBuffer(7, i))
'                                    End If
'                                    .DS.Fields("11").Updated = True
'                                    LogFieldValue .DS.Fields("11")
'
'                                    ' Imposta i litri caricati
'                                    .DS.Fields("7").NewValue = .DS.Fields("11").NewValue
'                                    .DS.Fields("7").Updated = True
'                                    LogFieldValue .DS.Fields("7")
'
'                                    ' Aggiorna l'ora fine scarico (VVPHRFIC)
'                                    .DS.Fields("39").NewValue = TruncDate(dateLocalEventDateTime)
'                                    .DS.Fields("39").Updated = True
'                                    LogFieldValue .DS.Fields("39")
'
'                                    ' Aggiorna il flag Modificato (VVPFLST5)
'                                    .DS.Field("Stato Riga 5").NewValue = "S"
'                                    .DS.Field("Stato Riga 5").Updated = True
'                                    LogFieldValue .DS.Field("Stato Riga 5")
'                                End If
'
'                                ' Salva
'                                If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
'                                    objLogger.logError "Save Consegna Prodotto."
'                                    GoTo ErrHandler
'                                End If
'
'                                .DS.DA.EndTrans etCommit
'                            End With
'
'                            ' Carica la consegna e aggiorna i dati
'                            With VVC
'
'                                .DS.Fields.ResetFields
'                                .DS.Fields("0").FindValue = varIDViaggio
'                                .DS.Fields("1").FindValue = varIDConsegna
'
'                                If Not .Find Then
'                                    objLogger.logError "Find Consegna."
'                                    GoTo ErrHandler
'                                End If
'
'                                ' Se e' l'ultimo scarico della consegna, aggiorna il tempo scarico e lo stato
'                                lngHowMany = 0
'                                If Not DS.DA.QueryCommand(SQL_KNOSOS_GET_COUNT_SCARICHI_NONCONCLUSI, varBuffer, lngHowMany, _
'                                    varIDViaggio, varIDConsegna) Then
'                                    objLogger.logError "Query SQL_KNOSOS_GET_COUNT_SCARICHI_NONCONCLUSI."
'                                    GoTo ErrHandler
'                                End If
'
'                                If CLng(varBuffer(0, 0)) = 0 Then
'
'                                    If blnUsaDatiGPS Then
'                                        ' Aggiorna il tempo di scarico
'                                        .DS.Fields("25").NewValue = DateDiff("n", TruncDate(.DS.Fields("24").Value), TruncDate(dateLocalEventDateTime))
'                                        .DS.Fields("25").Updated = True
'                                        LogFieldValue .DS.Fields("25")
'
'                                        ' Aggiorna il flag consegna ultimata
'                                        .DS.Fields("36").NewValue = "S"
'                                        .DS.Fields("36").Updated = True
'                                        LogFieldValue .DS.Fields("36")
'                                    End If
'
'                                    ' Aggiorna lo stato 'Completo'
'                                    .DS.Fields("44").NewValue = "2"
'                                    .DS.Fields("44").Updated = True
'                                    LogFieldValue .DS.Fields("44")
'                                End If
'
''                                ' Imposta i km parziali
''                                If Not (IsNull(varKBuffer(9, i)) Or IsEmpty(varKBuffer(9, i))) Then
''                                    .DS.Fields("22").NewValue = GetKmConsegna(varIDViaggio, varIDConsegna) + CSng(varKBuffer(9, i))
''                                    .DS.Fields("22").Updated = True
''                                End If
'
'                                ' Salva la consegna
'                                If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
'                                    objLogger.logError "Save Consegna."
'                                    GoTo ErrHandler
'                                End If
'
'                                .DS.DA.EndTrans etCommit
'                            End With
                        
                        ' -------------------------------------------------------
                        ' Evento - Fin de viaje (Fine del viaggio)
                        ' -------------------------------------------------------
                        Case "OFF"
'                            Dim sngImportoPedaggio As Single
'                            Dim lngKmInizioViaggio As Long
'                            Dim lngKmInizioCarico As Long
'
                            objLogger.logInfo "Tipo evento: Fine viaggio"

                            DS.Fields("4").NewValue = "VF"
                            DS.Fields("4").Updated = True
'
                            If CDbl(varKBuffer(5, i)) > 0 Then
                                ' Controllo di non avere altri eventi OFF in precedenza per questo stesso viaggio
                                varBuffer = Empty
                                If DS.DA.QueryCommand(SQL_GET_EVENTI_PRECEDENTI, varBuffer, 0, CVar(varKBuffer(4, i)), CVar(varKBuffer(2, i)), CVar(varKBuffer(7, i))) Then
                                    If IsEmpty(verBuffer) Then
                                        ' Aggiorno i Km di rientro reali per il viaggio in corso
                                        VVT.DS.Fields("57").NewValue = CDbl(varKBuffer(5, i)) / 1000
                                        VVT.DS.Fields("57").Updated = True
                                        LogFieldValue VVT.DS.Fields("57")
                                    Else
                                        objLogger.logWarning "Evento OFF ignorato in quanto ne è già stato ricevuto un altro in precedenza per lo stesso viaggio"
                                    End If
                                Else
                                    objLogger.logError "Errore durante il controllo sull'evento OFF"
                                End If
                            Else
                                objLogger.logWarning "Evento OFF ignorato in quanto giunto con i Km a 0"
                            End If
                            
                        
'                            With VVT
'                                ' Aggiorna lo stato del viaggio (VVTFLSTA)
'                                .DS.Fields("48").NewValue = VIAGGIO_STATO_CONCLUSO
'                                .DS.Fields("48").Updated = True
'                                LogFieldValue .DS.Fields("48")
'
'                                If blnUsaDatiGPS Then
'                                    ' Aggiorna l'ora fine viaggio (VVTHRFIV)
'                                    .DS.Fields("9").NewValue = TruncDate(dateLocalEventDateTime)
'                                    .DS.Fields("9").Updated = True
'                                    LogFieldValue .DS.Fields("9")
'
'                                    ' Aggiorna il tempo del viaggio sul viaggio (VVTTMPVI)
'                                    .DS.Fields("29").NewValue = DateDiff("n", TruncDate(.DS.Fields("7").Value), TruncDate(dateLocalEventDateTime))
'                                    .DS.Fields("29").Updated = True
'                                    LogFieldValue .DS.Fields("29")
'
'                                    ' Preleva i km di fine carico
'                                    lngHowMany = 0
'                                    If Not DS.DA.QueryCommand(SQL_GET_KM_EVENT, varBuffer, lngHowMany, varIDViaggio, "CF") Then
'                                        objLogger.logError "Query SQL_GET_KM_EVENT."
'                                        GoTo ErrHandler
'                                    End If
'
'                                    If lngHowMany = 0 Then
'                                        objLogger.logError "Impossibile calcolare il km definitivi del viaggio. Manca l'evento fine carico."
'                                    Else
'                                        ' Aggiorna i km definitivi sul viaggio
'                                        If Not (IsNull(varKBuffer(9, i)) Or IsEmpty(varKBuffer(9, i))) Then
'                                            .DS.Fields("27").NewValue = CDbl(varKBuffer(9, i)) - varBuffer(0, 0)
'                                            .DS.Fields("27").Updated = True
'                                        End If
'                                        LogFieldValue .DS.Fields("27")
'                                    End If
'
'                                    ' Preleva i km di inizio viaggio
'                                    lngHowMany = 0
'                                    If Not DS.DA.QueryCommand(SQL_GET_KM_EVENT, varBuffer, lngHowMany, varIDViaggio, "VI") Then
'                                        objLogger.logError "Query SQL_GET_KM_EVENT."
'                                        GoTo ErrHandler
'                                    End If
'
'                                    If lngHowMany = 0 Then
'                                        objLogger.logError "Impossibile aggiornare i km a vuoto. Manca l'evento inizio viaggio."
'                                    Else
'                                        lngKmInizioViaggio = varBuffer(0, 0)
'
'                                        ' Preleva i km di inizio carico
'                                        lngHowMany = 0
'                                        If Not DS.DA.QueryCommand(SQL_GET_KM_EVENT, varBuffer, lngHowMany, varIDViaggio, "CI") Then
'                                            objLogger.logError "Query SQL_GET_KM_EVENT."
'                                            GoTo ErrHandler
'                                        End If
'
'                                        If lngHowMany = 0 Then
'                                            objLogger.logError "Impossibile aggiornare i km a vuoto. Manca l'evento inizio carico."
'                                        Else
'                                            lngKmInizioCarico = varBuffer(0, 0)
'
'                                            ' Aggiorna i km a vuoto
'                                            .DS.Fields("49").NewValue = lngKmInizioCarico - lngKmInizioViaggio
'                                            .DS.Fields("49").Updated = True
'                                            LogFieldValue .DS.Fields("49")
'                                        End If
'                                    End If
'
'                                    ' Aggiorna il flag modificato (VVPFLST5) di tutti i prodotti del viaggio
'                                    If Not DS.DA.ExecuteCommand(SQL_UPDATE_VIAGGIO_MODIFICATO, varIDViaggio) Then
'                                        objLogger.logError "Errore in aggiornamento flag Modificato del viaggio."
'                                        GoTo ErrHandler
'                                    End If
'                                End If
'
'                                ' Aggiorna i km totali del viaggio
'                                .DS.Fields("49").NewValue = CLng(varKBuffer(9, i)) - lngKmInizioViaggio
'                                .DS.Fields("49").Updated = True
'                                LogFieldValue .DS.Fields("49")
'
'                            End With
'
'                            If blnUsaDatiGPS Then
'                                ' Preleva l'importo pedaggio
'                                lngHowMany = 0
'                                If Not DS.DA.QueryCommand(SQL_KNOSOS_GET_IMPORTO_PEDAGGIO, varKBuffer2, lngHowMany, _
'                                                CStr(varKBuffer(2, i)), KNOSOS_VIAGGIO_STATO_TERMINATO) Then
'                                    objLogger.logError "Query SQL_KNOSOS_GET_IMPORTO_PEDAGGIO."
'                                    GoTo ErrHandler
'                                End If
'
'                                If lngHowMany > 0 Then
'                                    sngImportoPedaggio = CSng(Replace(varKBuffer2(0, 0), ".", ","))
'                                    DS.DA.LogExecTime "INFO: - PEDAGGIO  [" & varKBuffer2(0, 0) & "] [" & sngImportoPedaggio & "]", -1
'                                Else
'                                    sngImportoPedaggio = 0
'                                    DS.DA.LogExecTime "INFO: - PEDAGGIO NON TROVATO", -1
'                                    blnAttesaPedaggio = True
'                                End If
'
'                                ' Calcola imponibile pedaggio
'                                If sngImportoPedaggio > 0 Then
'                                    sngImportoPedaggio = Round(sngImportoPedaggio / (1 + (ALIQUOTA_IVA_PEDAGGIO / 100)), 2)
'                                End If
'
'                                ' Aggiorna il campo importo pedaggio
'                                VVT.DS.Fields("46").NewValue = sngImportoPedaggio
'                                VVT.DS.Fields("46").Updated = True
'                                LogFieldValue VVT.DS.Fields("46")
'
'                                ' Aggiorna i km
'                                If Not UpdateKM(varIDViaggio, CSng(varKBuffer(9, i))) Then
'                                    GoTo ErrHandler
'                                End If
'                            End If

                        Case Else
                            objLogger.logError "Tipo evento sconosciuto."
                            GoTo ErrHandler

                    End Select

                    ' Aggiorna il viaggio
                    With VVT
'                        If Not (IsNull(varKBuffer(9, i)) Or IsEmpty(varKBuffer(9, i))) Then
'                            .DS.Fields("27").NewValue = GetKmViaggio(varIDViaggio) + CSng(varKBuffer(9, i))
'                            .DS.Fields("27").Updated = True
'                        End If
'
'                        If blnUsaDatiGPS Then
'                            ' Aggiorna il flag trasmissione viaggio
'                            .DS.Fields("41").NewValue = "S"
'                            .DS.Fields("41").Updated = True
'                            LogFieldValue .DS.Fields("41")
'                        End If

                        ' Salva
                        If Not .DS.BeginSave(.SQL_UPDATE, .SQL_GET_TS, False) Then
                            objLogger.logError "Save Viaggio."
                            GoTo ErrHandler
                        End If
        
                        .DS.DA.EndTrans etCommit
                    End With
                    
'                    ' Elimina l'evento KNOSOS
'                    objLogger.logInfo "Cancellazione KNOSOS EVENTO."
'                    strSQLStatement = Replace(SQL_KNOSOS_DELETE_EVENT, "%fecha%", varKBuffer(3, i))
'
'                    ' Se il pedaggio non è ancora stato inserito, rimandiamo l'eliminazione dell'evento OFF da KNOSOS
'                    ' e la creazione dell'evento Fine Viaggio in RESOil.2
'                    If Not blnAttesaPedaggio Then
'                        If Not DS.DA.ExecuteCommand(strSQLStatement, CStr(varKBuffer(0, i)), _
'                                                    CStr(varKBuffer(1, i)), _
'                                                    CStr(varKBuffer(2, i)), _
'                                                    CLng(varKBuffer(4, i))) Then
'                            objLogger.logError "Cancellazione Evento KNOSOS."
'                            GoTo ErrHandler
'                        End If
'
'                        objLogger.logInfo "Creazione dell'evento in ResOil."
'                        If Not Insert Then
'                            objLogger.logError "Insert ResOil evento."
'                            GoTo ErrHandler
'                        End If
'                    End If
                End If
            Else
                DS.DA.LogExecTime "WARN: Viaggio non trovato!", -1
            End If

NextIteration:
            DS.DA.EndTrans etCommit
            DoEvents
            LastID = CVar(varKBuffer(7, i))
        Next i
    End If

    ScanEvents_OESIA = True
       
    Set VVP = Nothing
    Set VVT = Nothing
    Set VVC = Nothing
    Set ADZ = Nothing

    objLogger.logDebug "Fine ScanEvents_OESIA senza errori."
    Exit Function

ErrHandler:
    DS.DA.EndTrans etRollback
    
    Set VVP = Nothing
    Set VVT = Nothing
    Set VVC = Nothing
    Set ADZ = Nothing
    
    objLogger.logError "Fine ScanEvents_OESIA con ERRORI."

    ScanEvents_OESIA = False
End Function

