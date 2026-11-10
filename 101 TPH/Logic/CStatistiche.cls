VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStatistiche"
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

Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1

Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)
Private mcolListFields As Collection

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
Private Sub Class_Initialize()
'    SQL_FIND = "SELECT ampidamp, ampdescr, ampannul, timestmp FROM anagamp0 WHERE "
'    SQL_GETLIST = "SELECT ampidamp, ampdescr FROM anagamp0 "
'    SQL_UPDATE = "UPDATE anagamp0 SET "
'    SQL_INSERT = "INSERT INTO anagamp0 "
'    SQL_GET_TS = "SELECT timestmp FROM anagamp0 WHERE "
    
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Dal Numero Viaggio", "viagvvt0.vvtidvvt", 0, , 7, "0", " >= "
        .Add "Al Numero Viaggio", "viagvvt0.vvtidvvt", 0, , 7, "1", " <= "
        .Add "Dalla Data Inizio Viaggio", "viagvvt0.vvtdtinv", 2, , 10, "5", " >= "
        .Add "Alla Data Inizio Viaggio", "viagvvt0.vvtdtinv", 2, , 10, "6", " <= "
        .Add "Deposito", "viagvvt0.vvtidade", 0, , 3, "7"
        With .Item("7")
            .ForeignTableEntity = "ADE"
            .ForeignTable = "anagade0"
            .ForeignTableFld = "adeidade"
            .ForeignTableDescField = "adedescr"
            .ForeignTableDeleteFld = "adeannul"
        End With
        .Add "Dal Numero Viaggio Definitivo", "viagvvt0.vvtiddef", 0, , 7, "11", " >= "
        .Add "Al Numero Viaggio Definitivo", "viagvvt0.vvtiddef", 0, , 7, "12", " <= "
    
        .Add "Dal Numero Viaggio", "prefppd0.ppdidvvt", 0, , 7, "13", " >= "
        .Add "Al Numero Viaggio", "prefppd0.ppdidvvt", 0, , 7, "14", " <= "
        .Add "Dalla Data Inizio Viaggio", "prefppd0.ppddtvia", 2, , 10, "15", " >= "
        .Add "Alla Data Inizio Viaggio", "prefppd0.ppddtvia", 2, , 10, "16", " <= "
        .Add "Deposito", "prefppd0.ppdidade", 0, , 3, "17"
        With .Item("17")
            .ForeignTableEntity = "ADE"
            .ForeignTable = "anagade0"
            .ForeignTableFld = "adeidade"
            .ForeignTableDescField = "adedescr"
            .ForeignTableDeleteFld = "adeannul"
        End With
        .Add "Zona Ordine", "ordioct0.octidazo", 0, , 3, "18"
        .Add "Dalla Data Ricezione Ordine", "ordioct0.octdtric", 2, , 10, "19", " >= "
        .Add "Alla Data Ricezione Ordine", "ordioct0.octdtric", 2, , 10, "20", " <= "
        .Add "Dalla Data Inserimento Ordine", "ordioct0.octdtins", 2, , 10, "21", " >= "
        .Add "Alla Data Inserimento Ordine", "ordioct0.octdtins", 2, , 10, "22", " <= "
    End With
    
    With DS.KeyFields
        .Add "0", "1"
        .Add "1", "2"
        .Add "5", "3"
        .Add "6", "4"
        .Add "7", "5"
        .Add "11", "6"
        .Add "12", "7"
        .Add "13", "8"
        .Add "14", "9"
        .Add "15", "10"
        .Add "16", "11"
        .Add "17", "12"
        .Add "18", "13"
    End With
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


