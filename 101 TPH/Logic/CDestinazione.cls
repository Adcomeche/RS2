VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDestinazione"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Public SQL_GETLIST As String
Public SQL_GETLIST_CLIENT As String
Public SQL_FIND As String
Public SQL_UPDATE As String
Public SQL_INSERT As String
Public SQL_GET_TS As String
Public SQL_GET_SEARCH_LIST As String
Public SQL_GET_SEARCH_LIST_GROUPBY As String
Public SQL_COUNT_CAPCISTERNE As String
Private SQL_GET_LIST_DISTANZE As String
Private SQL_UPDATE_DISTANZE As String
Private SQL_INSERT_DISTANZE As String
Private SQL_DELETE_DISTANZA As String
Private SQL_GET_DISTANZA As String
Private SQL_GET_SERVIZI_SPECIALI As String
Private SQL_GET_SELECTION_EXPORT As String

Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1
Private mcolListFields As Collection
Private mvarFieldsIndexes As Variant

Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)

Public Function DecodeFirmCode(FirmID As String, FirmCode As String, Optional ReturnID As Boolean = False) As String
    On Error GoTo ErrHandler
    
    DecodeFirmCode = DS.DecodeFirmCode(FirmID, FirmCode, Me.SQL_FIND, ReturnID)
    
    Exit Function
ErrHandler:
    DecodeFirmCode = vbNullString
End Function

Public Property Get SearchListFields() As Collection
    Set SearchListFields = mcolListFields
End Property

Public Function Find() As Boolean
    Find = DS.Find(Me.SQL_FIND)
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
Public Function UpdateDistanze(IDDestinazione As String, IDDeposito As String, KmAttivi As Long, KmPassivi As Long, Pedaggio As Single) As Boolean
    On Error GoTo ErrHandler
    
    If Not IsNumeric(KmAttivi) Then KmAttivi = 0
    If Not IsNumeric(KmPassivi) Then KmPassivi = 0
    If Not IsNumeric(Pedaggio) Then Pedaggio = 0
    
    UpdateDistanze = DS.DA.ExecuteCommand(SQL_UPDATE_DISTANZE, KmAttivi, KmPassivi, Pedaggio, IDDestinazione, IDDeposito)
    
    Exit Function
ErrHandler:
    UpdateDistanze = False
End Function

Public Function InsertDistanze(IDDestinazione As String, IDDeposito As String, KmAttivi As Long, KmPassivi As Long, Pedaggio As Single) As Boolean
    On Error GoTo ErrHandler
    
    If Not IsNumeric(KmAttivi) Then KmAttivi = 0
    If Not IsNumeric(KmPassivi) Then KmPassivi = 0
    If Not IsNumeric(Pedaggio) Then Pedaggio = 0
    
    InsertDistanze = DS.DA.ExecuteCommand(SQL_INSERT_DISTANZE, IDDestinazione, IDDeposito, KmAttivi, KmPassivi, Pedaggio)
    
    Exit Function
ErrHandler:
    InsertDistanze = False
End Function

Public Function DeleteDistanze(IDDestinazione As String, IDDeposito As String) As Boolean
    On Error GoTo ErrHandler
    
    DeleteDistanze = DS.DA.ExecuteCommand(SQL_DELETE_DISTANZA, IDDestinazione, IDDeposito)
    
    Exit Function
ErrHandler:
    DeleteDistanze = False
End Function


Private Sub Class_Initialize()
    Dim i As Integer
    
'    #If (VersionID = 4) Or (VersionID = 5) Then
'        SQL_FIND = "SELECT adzidadz, adzrasoc, adzrasag, adzrasab, " & _
'                   "adzindir, adzlocal, adzprovi, adznrcap, adzidacl, " & _
'                   "adzcofis, adzpaiva, adzidard, adzidatd, adzidamp, " & _
'                   "adzidaiv, adzidade, adzidazo, adzcooca, adznrtel, " & _
'                   "adznota1, adznota2, adzragu1, adzragu2, adzragu3, " & _
'                   "adzragu4, adzragu5, adzidaee, adzidatv, adzannul, timestmp FROM anagadz0 WHERE "
'    #Else
        SQL_FIND = "SELECT adzidadz, adzrasoc, adzrasag, adzrasab, " & _
                   "adzindir, adzlocal, adzprovi, adznrcap, adzidacl, " & _
                   "adzcofis, adzpaiva, adzidard, adzidatd, adzidamp, " & _
                   "adzidaiv, adzidade, adzidazo, adzcooca, adznrtel, " & _
                   "adznota1, adznota2, adzragu1, adzragu2, adzragu3, " & _
                   "adzragu4, adzragu5, adzkmrag, adznripa, adzcgpsx, " & _
                   "adzcgpsy, adzhrap1, adzhrch1, adzhrap2, adzhrch2, " & _
                   "adzhrap3, adzhrch3, adznrte1, adznrte2, adznrtem, " & _
                   "adzannul, timestmp FROM anagadz0 WHERE "
'    #End If
    SQL_GETLIST = "SELECT adzidadz, adzrasoc, adzindir, adzlocal, adzprovi, adzannul FROM anagadz0 "
    SQL_GETLIST_CLIENT = "SELECT adzidadz, adzrasoc, adzindir, adzlocal, adzprovi, adzannul FROM anagadz0 WHERE adzidacl = "
    SQL_UPDATE = "UPDATE anagadz0 SET "
    SQL_INSERT = "INSERT INTO anagadz0 "
    SQL_GET_TS = "SELECT timestmp FROM anagadz0 WHERE "
    SQL_GET_SEARCH_LIST = "SELECT adzidadz, adzrasoc, adzidacl, adzindir, adzlocal, adzprovi, adznrcap, adzannul FROM anagadz0 "
    SQL_COUNT_CAPCISTERNE = "SELECT SUM(APZCACIS) as TOTCIS FROM prezapz0 WHERE apzidadz=? AND apzidapr=?"
    SQL_GET_LIST_DISTANZE = "SELECT DEDIDADE, DEDKMATT, DEDKMPAS, DEDNRIPA " & _
                            "  FROM ANAGDED0 " & _
                            " WHERE DEDIDADZ = ? "
    SQL_UPDATE_DISTANZE = "UPDATE ANAGDED0 " & _
                          "   SET DEDKMATT = ? , DEDKMPAS = ? , DEDNRIPA = ? " & _
                          " WHERE DEDIDADZ = ? AND DEDIDADE = ? "
    SQL_INSERT_DISTANZE = "INSERT INTO ANAGDED0 (DEDIDADZ, DEDIDADE, DEDKMATT, DEDKMPAS, DEDNRIPA)" & _
                          " VALUES (?, ?, ?, ?, ?) "
    SQL_GET_DISTANZA = "SELECT DEDKMATT, DEDKMPAS, DEDNRIPA " & _
                       "  FROM ANAGDED0 " & _
                       " WHERE DEDIDADZ = ? AND DEDIDADE = ? "
    
    SQL_DELETE_DISTANZA = "DELETE ANAGDED0 WHERE DEDIDADZ = ? AND DEDIDADE = ?"
    
    SQL_GET_SELECTION_EXPORT = "SELECT ADZIDADZ, ADZRASOC, ADZRASAG, ADZINDIR, ADZLOCAL, ADZRAGU4" & _
                        " FROM ANAGADZ0"
    
    SQL_GET_SERVIZI_SPECIALI = "SELECT SXDFLTSE, SXDNRALQ, SXDIDAUM " & _
                               "  FROM TARISXD0 " & _
                               " WHERE SXDIDADZ = ? AND SXDDTDAT <= ? " & _
                               " ORDER BY SXDDTDAT DESC "
    
    Set DS = New CDataStruct
    Set DS.MultiSoc = New CDestinazioneSoc
    
    mvarFieldsIndexes = Array("0", "1", "8", "4", "5", "6", "7", "26")
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i
    
    With DS.Fields
        .Add "Codice", "anagadz0.adzidadz", 0, True, 10, "0"
        .Add "Ragione Sociale", "anagadz0.adzrasoc", 0, True, 40, "1"
        .Add "Ragione Sociale Aggiuntiva", "anagadz0.adzrasag", 0, False, 40, "2"
        .Add "Ragione Sociale Abbreviata", "anagadz0.adzrasab", 0, False, 6, "3"
        .Add "Indirizzo", "anagadz0.adzindir", 0, False, 40, "4"
        .Add "Località", "anagadz0.adzlocal", 0, False, 25, "5"
        .Add "Provincia", "anagadz0.adzprovi", 0, False, 2, "6"
        .Add "CAP", "anagadz0.adznrcap", 0, False, 10, "7"
        .Add "Cliente", "anagadz0.adzidacl", 0, True, 10, "8"
        With .Item("8")
            .ForeignTableEntity = "ACL"
            .ForeignTable = "anagacl0"
            .ForeignTableFld = "aclidacl"
            .ForeignTableDescField = "aclrasoc"
            .ForeignTableDeleteFld = "aclannul"
        End With
        .Add "Codice Fiscale", "anagadz0.adzcofis", 0, False, 16, "9"
        .Add "Partita IVA", "anagadz0.adzpaiva", 0, False, 20, "10"
        .Add "Raggruppamento Destinazione", "anagadz0.adzidard", 0, False, 3, "11"
        With .Item("11")
            .ForeignTableEntity = "ARD"
            .ForeignTable = "anagard0"
            .ForeignTableFld = "ardidard"
            .ForeignTableDescField = "arddescr"
            .ForeignTableDeleteFld = "ardannul"
        End With
        .Add "Tipo Destinazione", "anagadz0.adzidatd", 0, True, 3, "12"
        With .Item("12")
            .ForeignTableEntity = "ATD"
            .ForeignTable = "anagatd0"
            .ForeignTableFld = "atdidatd"
            .ForeignTableDescField = "atddescr"
            .ForeignTableDeleteFld = "atdannul"
        End With
        .Add "Modalità di Pagamento", "anagadz0.adzidamp", 0, False, 3, "13"
        With .Item("13")
            .ForeignTableEntity = "AMP"
            .ForeignTable = "anagamp0"
            .ForeignTableFld = "ampidamp"
            .ForeignTableDescField = "ampdescr"
            .ForeignTableDeleteFld = "ampannul"
        End With
        .Add "IVA", "anagadz0.adzidaiv", 0, False, 3, "14"
        With .Item("14")
            .ForeignTableEntity = "AIV"
            .ForeignTable = "anagaiv0"
            .ForeignTableFld = "aividaiv"
            .ForeignTableDescField = "aivdescr"
            .ForeignTableDeleteFld = "aivannul"
        End With
        .Add "Deposito", "anagadz0.adzidade", 0, True, 3, "15"
        With .Item("15")
            .ForeignTableEntity = "ADE"
            .ForeignTable = "anagade0"
            .ForeignTableFld = "adeidade"
            .ForeignTableDescField = "adedescr"
            .ForeignTableDeleteFld = "adeannul"
        End With
        .Add "Zona", "anagadz0.adzidazo", 0, True, 3, "16"
        With .Item("16")
            .ForeignTableEntity = "AZO"
            .ForeignTable = "anagazo0"
            .ForeignTableFld = "azoidazo"
            .ForeignTableDescField = "azodescr"
            .ForeignTableDeleteFld = "azoannul"
        End With
        .Add "Coordinate Cartografia", "anagadz0.adzcooca", 0, False, 10, "17"
        .Add "Telefono", "anagadz0.adznrtel", 0, False, 20, "18"
        .Add "Nota 1", "anagadz0.adznota1", 0, False, 40, "19"
        .Add "Nota 2", "anagadz0.adznota2", 0, False, 40, "20"
        
        ' Pedidos Eventuales
        .Add "Raggruppamento Utente 1", "anagadz0.adzragu1", 0, False, 2, "21"
        
        .Add "Raggruppamento Utente 2", "anagadz0.adzragu2", 0, False, 2, "22"
        .Add "Raggruppamento Utente 3", "anagadz0.adzragu3", 0, False, 2, "23"
        .Add "Raggruppamento Utente 4", "anagadz0.adzragu4", 0, False, 2, "24"
        
        ' Raggruppamento geografico per stipendi
        .Add "Raggruppamento Utente 5", "anagadz0.adzragu5", 0, False, 2, "25"
        
        .Add "Km Raggio", "anagadz0.adzkmrag", 1, False, 6, "29"
        .Add "Importo pedaggio", "anagadz0.adznripa", 1, False, 13, "30"
        .Add "Coordinate GPS X", "anagadz0.adzcgpsx", 0, False, 15, "31"
        .Add "Coordinate GPS Y", "anagadz0.adzcgpsy", 0, False, 15, "32"
        .Add "Ora Apertura 1", "anagadz0.adzhrap1", 4, False, 5, "33"
        .Add "Ora Chiusura 1", "anagadz0.adzhrch1", 4, False, 5, "34"
        .Add "Ora Apertura 2", "anagadz0.adzhrap2", 4, False, 5, "35"
        .Add "Ora Chiusura 2", "anagadz0.adzhrch2", 4, False, 5, "36"
        .Add "Ora Apertura 3", "anagadz0.adzhrap3", 4, False, 5, "37"
        .Add "Ora Chiusura 3", "anagadz0.adzhrch3", 4, False, 5, "38"
        
        '-- DP 14/08/08
        .Add "Telefono 1", "anagadz0.adznrte1", 0, False, 20, "39"
        .Add "Telefono 2", "anagadz0.adznrte2", 0, False, 20, "40"
        .Add "Telefono Movil", "anagadz0.adznrtem", 0, False, 20, "41"
        
        .Add "Annullato", "anagadz0.adzannul", 3, , , "26"
    End With
    With DS.KeyFields
        .Add "0", "1"
    End With
    DS.CancelField = "anagadz0.adzannul"
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
Public Function ContaCapacitaCisterne(CodADZ As String, CodAPR As String) As Integer
    On Error GoTo ErrHandler
    Dim varBuffer As Variant
    Dim blnResult As Boolean
    
    blnResult = DS.DA.QueryCommand(SQL_COUNT_CAPCISTERNE, varBuffer, _
                                   0, CodADZ, CodAPR)
    If Not blnResult Then
        GoTo ErrHandler
    End If
    ContaCapacitaCisterne = varBuffer(0, 0)
    
    Exit Function
ErrHandler:
    ContaCapacitaCisterne = -1
End Function


Public Function GetListDistanze(IDDestinazione As String) As Variant
    On Error GoTo ErrHandler
    Dim varBuffer As Variant
    Dim blnResult As Boolean
    
    blnResult = DS.DA.QueryCommand(SQL_GET_LIST_DISTANZE, varBuffer, _
                                   0, IDDestinazione)
    If Not blnResult Then
        GoTo ErrHandler
    End If
    GetListDistanze = varBuffer
    
    Exit Function
ErrHandler:
    GetListDistanze = Empty
End Function

Public Function GetDistanza(IDDestinazione As String, IDDeposito As String) As Variant
    On Error GoTo ErrHandler
    Dim varBuffer As Variant
    Dim blnResult As Boolean
    
    blnResult = DS.DA.QueryCommand(SQL_GET_DISTANZA, varBuffer, _
                                   0, IDDestinazione, IDDeposito)
    If Not blnResult Then
        GoTo ErrHandler
    End If
    GetDistanza = varBuffer
    
    Exit Function
ErrHandler:
    GetDistanza = Empty
End Function

Public Function GetServiziSpeciali(IDDestinazione As String, Data As Variant) As Variant
    On Error GoTo ErrHandler
    Dim varBuffer As Variant
    Dim blnResult As Boolean
    
    blnResult = DS.DA.QueryCommand(SQL_GET_SERVIZI_SPECIALI, varBuffer, _
                                   0, IDDestinazione, Data)
    If Not blnResult Then
        GoTo ErrHandler
    End If
    GetServiziSpeciali = varBuffer
    
    Exit Function
ErrHandler:
    GetServiziSpeciali = Empty
End Function


Public Function GetSelectionExport(Optional ActiveRecordsOnly As Boolean = False, _
                             Optional OrderFields As Variant, _
                             Optional MaxRecords As Long = 0) As Boolean
    If Not IsMissing(OrderFields) Then
        GetSelectionExport = DS.GetSelection(SQL_GET_SELECTION_EXPORT, ActiveRecordsOnly, OrderFields, MaxRecords)
    Else
        GetSelectionExport = DS.GetSelection(SQL_GET_SELECTION_EXPORT, ActiveRecordsOnly, , MaxRecords)
    End If

End Function


