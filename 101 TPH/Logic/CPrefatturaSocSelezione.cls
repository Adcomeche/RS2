VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPrefatturaSocSelezione"
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
Public SQL_GET_SELECTION As String
Public SQL_GET_SELECTION_CHECK As String
Public SQL_GET_SELECTION_PEDAGGIO As String

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
    SQL_FIND = "SELECT gp1idazo, gp1idade, gp1rife1, gp1dtcon, " & _
               "       gp1bosca, gp1qtsca, gp1rife2, gp1tarap, " & _
               "       gp1impor, gp1idapr, gp1idadz, gp1rasde, " & _
               "       gp1idaut, gp1qkmra, gp1flimf, gp1rife3, " & _
               "       gp1annoe, gp1idasc, gp1annul, timestmp FROM prefgp10 WHERE "
    SQL_GETLIST = "" '"SELECT ampidamp, ampdescr FROM prefgp10 "
    SQL_UPDATE = "UPDATE prefgp10 SET "
    SQL_INSERT = "INSERT INTO prefgp10 "
    SQL_GET_TS = "SELECT timestmp FROM prefgp10 WHERE "
    SQL_GET_SELECTION = "SELECT gp1annoe, gp1rife3, gp1idazo, gp1idade, " & _
                        "       gp1rife1, gp1dtcon, gp1bosca, gp1qtsca, " & _
                        "       gp1rife2, gp1tarap, gp1impor, gp1idapr, " & _
                        "       gp1idadz, gp1rasde, gp1idaut, gp1qkmra, " & _
                        "       gp1flimf, gp1annul " & _
                        "  FROM prefgp10 "
                        
    '-- DP 20/08/08
    '-- Aggiunta la colonna gp1imsex per l'importo servizi esenti
    SQL_GET_SELECTION_CHECK = "SELECT gp1annoe, gp1rife3, gp1idazo, gp1idade, " & _
                              "       gp1rife1, gp1dtcon, gp1bosca, gp1qtsca, " & _
                              "       gp1rife2, gp1tarap, gp1impor, gp1idapr, " & _
                              "       gp1idadz, gp1rasde, gp1idaut, gp1qkmra, " & _
                              "       gp1flimf, gp1annul, gp1riave, gp1riam1, " & _
                              "       gp1riam2, gp1ridut, gp1idau1, gp1niaau, " & _
                              "       gp1nrvia, gp1imppe, gp1qtskg, gp1impoe, " & _
                              "       gp1impbe, gp1imphe, gp1impse, gp1imsex " & _
                              "  FROM prefgp10 "
    SQL_GET_SELECTION_PEDAGGIO = "SELECT gp1imppe " & _
                                 "  FROM prefgp10 "
    
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Zona", "prefgp10.gp1idazo", 0, , 3, "0"
        With .Item("0")
            .ForeignTableEntity = "AZO"
            .ForeignTable = "anagazo0"
            .ForeignTableFld = "azoidazo"
            .ForeignTableDescField = "azodescr"
            .ForeignTableDeleteFld = "azoannul"
        End With
        .Add "Deposito", "prefgp10.gp1idade", 0, , 3, "1"
        With .Item("1")
            .ForeignTableEntity = "ADE"
            .ForeignTable = "anagade0"
            .ForeignTableFld = "adeidade"
            .ForeignTableDescField = "adedescr"
            .ForeignTableDeleteFld = "adeannul"
        End With
        .Add "Dal Riferimento AGIP", "prefgp10.gp1rife1", 0, , 15, "2", " >= "
        .Add "Al Riferimento AGIP", "prefgp10.gp1rife1", 0, , 15, "7", " <= "
        .Add "Riferimento Esterno", "prefgp10.gp1rife1", 0, , 15, "13"
        .Add "Data Consegna", "prefgp10.gp1dtcon", 2, , 10, "14", " = "
        .Add "Dalla Data Consegna", "prefgp10.gp1dtcon", 2, , 10, "3", " >= "
        .Add "Alla Data Consegna", "prefgp10.gp1dtcon", 2, , 10, "5", " <= "
        .Add "Dalla Bolletta di scarico", "prefgp10.gp1bosca", 0, , 15, "4", " >= "
        .Add "Alla Bolletta di scarico", "prefgp10.gp1bosca", 0, , 15, "6", " <= "
        .Add "Prodotto", "prefgp10.gp1idapr", 0, , 6, "9"
        With .Item("9")
            .ForeignTableEntity = "APR"
            .ForeignTable = "anagapr0"
            .ForeignTableFld = "apridapr"
            .ForeignTableDescField = "aprdesco"
            .ForeignTableDeleteFld = "aprannul"
        End With
        .Add "Destinazione", "prefgp10.gp1idadz", 0, , 10, "10"
        With .Item("10")
            .ForeignTableEntity = "ADZ"
            .ForeignTable = "anagadz0"
            .ForeignTableFld = "adzidadz"
            .ForeignTableDescField = "adzrasoc"
            .ForeignTableDeleteFld = "adzannul"
        End With
        .Add "Unità di Trasporto", "prefgp10.gp1idaut", 0, , 3, "12"
        With .Item("12")
            .ForeignTableEntity = "AUT"
            .ForeignTable = "anagaut0"
            .ForeignTableFld = "autidaut"
            .ForeignTableDescField = "autdescr"
            .ForeignTableDeleteFld = "autannul"
        End With
        .Add "Riferimento esterno 3", "prefgp10.gp1rife3", 0, , 15, "15"
        .Add "Anno prefattura", "prefgp10.gp1annoe", 1, , 4, "16"
        .Add "Società", "prefgp10.gp1idasc", 0, True, 3, "18"
        With .Item("18")
            .ForeignTableEntity = "ASC"
            .ForeignTable = "anagasc0"
            .ForeignTableFld = "ascidasc"
            .ForeignTableDescField = "ascrasoc"
            .ForeignTableDeleteFld = "ascannul"
        End With
        .Add "Ente Emittente", "prefgp10.gp1idaee", 0, True, 3, "47"
        With .Item("47")
            .ForeignTableEntity = "AEE"
            .ForeignTable = "anagaee0"
            .ForeignTableFld = "aeeidaee"
            .ForeignTableDescField = "aeedescr"
            .ForeignTableDeleteFld = "aeeannul"
        End With
    End With
    With DS.KeyFields
        .Add "0", "1"
        .Add "1", "2"
        .Add "2", "3"
        .Add "3", "4"
        .Add "4", "5"
        .Add "5", "6"
        .Add "6", "7"
        .Add "7", "8"
        .Add "9", "10"
        .Add "10", "11"
        .Add "12", "13"
        .Add "13", "14"
        .Add "14", "15"
        .Add "15", "16"
        .Add "16", "17"
        .Add "18", "18"
        .Add "47", "19"
    End With
    DS.CancelField = "prefgp10.gp1annul"
End Sub


Public Function GetSelection(Optional ActiveRecordsOnly As Boolean = False, _
                             Optional OrderFields As Variant, _
                             Optional MaxRecords As Long = 0) As Boolean
    
    If Not IsMissing(OrderFields) Then
        GetSelection = DS.GetSelection(SQL_GET_SELECTION, ActiveRecordsOnly, OrderFields, MaxRecords)
    Else
        GetSelection = DS.GetSelection(SQL_GET_SELECTION, ActiveRecordsOnly, , MaxRecords)
    End If
End Function


Public Function GetSelectionCheck(Optional ActiveRecordsOnly As Boolean = False, _
                                  Optional OrderFields As Variant, _
                                  Optional MaxRecords As Long = 0) As Boolean
    
    If Not IsMissing(OrderFields) Then
        GetSelectionCheck = DS.GetSelection(SQL_GET_SELECTION_CHECK, ActiveRecordsOnly, OrderFields, MaxRecords)
    Else
        GetSelectionCheck = DS.GetSelection(SQL_GET_SELECTION_CHECK, ActiveRecordsOnly, , MaxRecords)
    End If
End Function

Public Function GetSelectionPedaggio() As Boolean
    GetSelectionPedaggio = DS.GetSelection(SQL_GET_SELECTION_PEDAGGIO, True, , 0)
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


