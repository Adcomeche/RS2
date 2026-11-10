VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCliente"
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
Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1
Private mcolListFields As Collection
Private mvarFieldsIndexes As Variant
Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)


Public Function Find() As Boolean
    Find = DS.Find(Me.SQL_FIND)
End Function

Public Property Get SearchListFields() As Collection
    Set SearchListFields = mcolListFields
End Property
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
    
    SQL_FIND = "SELECT    aclidacl, aclrasoc, aclrasag, aclrasab, " & _
               "aclindir, acllocal, aclprovi, aclnrcap, aclidasc, " & _
               "aclidaee, aclcofis, aclpaiva, aclidarc, aclidatc, " & _
               "aclidamp, aclidaiv, aclnrtel, aclnota1, aclnota2, " & _
               "aclannul, timestmp FROM anagacl0 WHERE "
    SQL_GETLIST = "SELECT aclidamp, aclrasoc FROM anagacl0 "
    SQL_UPDATE = "UPDATE anagacl0 SET "
    SQL_INSERT = "INSERT INTO anagacl0 "
    SQL_GET_TS = "SELECT timestmp FROM anagacl0 WHERE "
    SQL_GET_SEARCH_LIST = "SELECT aclidacl, aclrasoc, " & _
               "aclindir, acllocal, aclprovi, aclnrcap, " & _
               "aclannul FROM anagacl0 "
    
    mvarFieldsIndexes = Array("0", "1", "4", "5", "6", "7", "19")
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i
    Set DS = New CDataStruct
    Set DS.MultiSoc = New CClienteSoc
    
    With DS.Fields
        .Add "Codice", "anagacl0.aclidacl", 0, True, 10, "0"
        .Add "Ragione Sociale", "anagacl0.aclrasoc", 0, True, 40, "1"
        .Add "Ragione Sociale Aggiuntiva", "anagacl0.aclrasag", 0, False, 40, "2"
        .Add "Ragione Sociale Abbreviata", "anagacl0.aclrasab", 0, False, 6, "3"
        .Add "Indirizzo", "anagacl0.aclindir", 0, False, 40, "4"
        .Add "Località", "anagacl0.acllocal", 0, False, 25, "5"
        .Add "Provincia", "anagacl0.aclprovi", 0, False, 2, "6"
        .Add "CAP", "anagacl0.aclnrcap", 0, False, 10, "7"
        .Add "Società", "anagacl0.aclidasc", 0, True, 3, "8"
        With .Item("8")
            .ForeignTableEntity = "ASC"
            .ForeignTable = "anagasc0"
            .ForeignTableFld = "ascidasc"
            .ForeignTableDescField = "ascrasoc"
            .ForeignTableDeleteFld = "ascannul"
        End With
        .Add "Ente Emittente", "anagacl0.aclidaee", 0, True, 3, "9"
        With .Item("9")
            .ForeignTableEntity = "AEE"
            .ForeignTable = "anagaee0"
            .ForeignTableFld = "aeeidaee"
            .ForeignTableDescField = "aeedescr"
            .ForeignTableDeleteFld = "aeeannul"
        End With
        .Add "Codice Fiscale", "anagacl0.aclcofis", 0, False, 16, "10"
        .Add "Partita IVA", "anagacl0.aclpaiva", 0, False, 20, "11"
        .Add "Raggruppamento Cliente", "anagacl0.aclidarc", 0, True, 3, "12"
        With .Item("12")
            .ForeignTableEntity = "ARC"
            .ForeignTable = "anagarc0"
            .ForeignTableFld = "arcidarc"
            .ForeignTableDescField = "arcdescr"
            .ForeignTableDeleteFld = "arcannul"
        End With
        .Add "Tipo Cliente", "anagacl0.aclidatc", 0, True, 3, "13"
        With .Item("13")
            .ForeignTableEntity = "ATC"
            .ForeignTable = "anagatc0"
            .ForeignTableFld = "atcidatc"
            .ForeignTableDescField = "atcdescr"
            .ForeignTableDeleteFld = "atcannul"
        End With
        .Add "Modalità di Pagamento", "anagacl0.aclidamp", 0, True, 3, "14"
        With .Item("14")
            .ForeignTableEntity = "AMP"
            .ForeignTable = "anagamp0"
            .ForeignTableFld = "ampidamp"
            .ForeignTableDescField = "ampdescr"
            .ForeignTableDeleteFld = "ampannul"
        End With
        .Add "IVA", "anagacl0.aclidaiv", 0, True, 3, "15"
        With .Item("15")
            .ForeignTableEntity = "AIV"
            .ForeignTable = "anagaiv0"
            .ForeignTableFld = "aividaiv"
            .ForeignTableDescField = "aivdescr"
            .ForeignTableDeleteFld = "aivannul"
        End With
        .Add "Telefono", "anagacl0.aclnrtel", 0, False, 20, "16"
        .Add "Nota 1", "anagacl0.aclnota1", 0, False, 40, "17"
        .Add "Nota 2", "anagacl0.aclnota2", 0, False, 40, "18"
        .Add "Annullato", "anagacl0.aclannul", 3, , , "19"
    End With
    With DS.KeyFields
        .Add "0", "1"
    End With
    DS.CancelField = "anagacl0.aclannul"
End Sub


Private Sub Class_Terminate()
    Set DS = Nothing
    Set mcolListFields = Nothing
End Sub


Private Sub DS_ClientClear()
    RaiseEvent ClientClear

End Sub

Private Sub DS_ClientRefresh()
    RaiseEvent ClientRefresh

End Sub




