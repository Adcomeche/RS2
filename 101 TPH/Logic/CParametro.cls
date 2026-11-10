VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CParametro"
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
    SQL_FIND = "SELECT axxidaxx, axxidaim, axxidapd, axxidacl, " & _
               "       axxidasc, axxidins, axxidp01, axxidp02, " & _
               "       axxidp03, axxidp04, axxidp05, axxidp06, " & _
               "       axxidp07, axxidp08, axxidp09, axxidp10, " & _
               "       axxidaee, axxidava, axxflsca, timestmp FROM anagaxx0 WHERE "
    SQL_GETLIST = "SELECT axxidaxx, axxidins FROM anagaxx0 "
    SQL_UPDATE = "UPDATE anagaxx0 SET "
    SQL_INSERT = "INSERT INTO anagaxx0 "
    SQL_GET_TS = "SELECT timestmp FROM anagaxx0 WHERE "
    
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Codice", "anagaxx0.axxidaxx", 0, True, 3, "0"
        .Add "Imballo Default", "anagaxx0.axxidaim", 0, False, 3, "1"
        With .Item("1")
            .ForeignTableEntity = "AIM"
            .ForeignTable = "anagaim0"
            .ForeignTableFld = "aimidaim"
            .ForeignTableDescField = "aimdescr"
            .ForeignTableDeleteFld = "aimannul"
        End With
        .Add "Pos. Doganale Default", "anagaxx0.axxidapd", 0, False, 3, "2"
        With .Item("2")
            .ForeignTableEntity = "APD"
            .ForeignTable = "anagapd0"
            .ForeignTableFld = "apdidapd"
            .ForeignTableDescField = "apddescr"
            .ForeignTableDeleteFld = "apdannul"
        End With
        .Add "Cliente Default", "anagaxx0.axxidacl", 0, False, 10, "3"
        With .Item("3")
            .ForeignTableEntity = "ACL"
            .ForeignTable = "anagacl0"
            .ForeignTableFld = "aclidacl"
            .ForeignTableDescField = "aclrasoc"
            .ForeignTableDeleteFld = "aclannul"
        End With
        .Add "Società Default", "anagaxx0.axxidasc", 0, True, 3, "4"
        With .Item("4")
            .ForeignTableEntity = "ASC"
            .ForeignTable = "anagasc0"
            .ForeignTableFld = "ascidasc"
            .ForeignTableDescField = "ascrasoc"
            .ForeignTableDeleteFld = "ascannul"
        End With
        .Add "Codice Installazione", "anagaxx0.axxidins", 0, True, 3, "5", , True
        .Add "Prodotto 1", "anagaxx0.axxidp01", 0, False, 6, "6"
        With .Item("6")
            .ForeignTableEntity = "APR"
            .ForeignTable = "anagAPR0"
            .ForeignTableFld = "APRidAPR"
            .ForeignTableDescField = "aprdesco"
            .ForeignTableDeleteFld = "APRannul"
        End With
        .Add "Prodotto 2", "anagaxx0.axxidp02", 0, False, 6, "7"
        With .Item("7")
            .ForeignTableEntity = "APR"
            .ForeignTable = "anagAPR0"
            .ForeignTableFld = "APRidAPR"
            .ForeignTableDescField = "aprdesco"
            .ForeignTableDeleteFld = "APRannul"
        End With
        .Add "Prodotto 3", "anagaxx0.axxidp03", 0, False, 6, "8"
        With .Item("8")
            .ForeignTableEntity = "APR"
            .ForeignTable = "anagAPR0"
            .ForeignTableFld = "APRidAPR"
            .ForeignTableDescField = "aprdesco"
            .ForeignTableDeleteFld = "APRannul"
        End With
        .Add "Prodotto 4", "anagaxx0.axxidp04", 0, False, 6, "9"
        With .Item("9")
            .ForeignTableEntity = "APR"
            .ForeignTable = "anagAPR0"
            .ForeignTableFld = "APRidAPR"
            .ForeignTableDescField = "aprdesco"
            .ForeignTableDeleteFld = "APRannul"
        End With
        .Add "Prodotto 5", "anagaxx0.axxidp05", 0, False, 6, "10"
        With .Item("10")
            .ForeignTableEntity = "APR"
            .ForeignTable = "anagAPR0"
            .ForeignTableFld = "APRidAPR"
            .ForeignTableDescField = "aprdesco"
            .ForeignTableDeleteFld = "APRannul"
        End With
        .Add "Prodotto 6", "anagaxx0.axxidp06", 0, False, 6, "11"
        With .Item("11")
            .ForeignTableEntity = "APR"
            .ForeignTable = "anagAPR0"
            .ForeignTableFld = "APRidAPR"
            .ForeignTableDescField = "aprdesco"
            .ForeignTableDeleteFld = "APRannul"
        End With
        .Add "Prodotto 7", "anagaxx0.axxidp07", 0, False, 6, "12"
        With .Item("12")
            .ForeignTableEntity = "APR"
            .ForeignTable = "anagAPR0"
            .ForeignTableFld = "APRidAPR"
            .ForeignTableDescField = "aprdesco"
            .ForeignTableDeleteFld = "APRannul"
        End With
        .Add "Prodotto 8", "anagaxx0.axxidp08", 0, False, 6, "13"
        With .Item("13")
            .ForeignTableEntity = "APR"
            .ForeignTable = "anagAPR0"
            .ForeignTableFld = "APRidAPR"
            .ForeignTableDescField = "aprdesco"
            .ForeignTableDeleteFld = "APRannul"
        End With
        .Add "Prodotto 9", "anagaxx0.axxidp09", 0, False, 6, "14"
        With .Item("14")
            .ForeignTableEntity = "APR"
            .ForeignTable = "anagAPR0"
            .ForeignTableFld = "APRidAPR"
            .ForeignTableDescField = "aprdesco"
            .ForeignTableDeleteFld = "APRannul"
        End With
        .Add "Prodotto 10", "anagaxx0.axxidp10", 0, False, 6, "15"
        With .Item("15")
            .ForeignTableEntity = "APR"
            .ForeignTable = "anagAPR0"
            .ForeignTableFld = "APRidAPR"
            .ForeignTableDescField = "aprdesco"
            .ForeignTableDeleteFld = "APRannul"
        End With
        .Add "Ente Emittente", "anagaxx0.axxidaee", 0, True, 3, "16"
        With .Item("16")
            .ForeignTableEntity = "AEE"
            .ForeignTable = "anagaee0"
            .ForeignTableFld = "aeeidaee"
            .ForeignTableDescField = "aeedescr"
            .ForeignTableDeleteFld = "aeeannul"
        End With
        .Add "Valuta", "anagaxx0.axxidava", 0, True, 3, "17"
        With .Item("17")
            .ForeignTableEntity = "AVA"
            .ForeignTable = "anagava0"
            .ForeignTableFld = "avaidava"
            .ForeignTableDescField = "avadescr"
            .ForeignTableDeleteFld = "avaannul"
        End With
        .Add "Controllo Scadenze", "anagaxx0.axxflsca", 0, True, 2, "18"
    End With
    With DS.KeyFields
        .Add "0", "1"
    End With
    DS.CancelField = ""
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




