VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRigaDocumentoAmministrativo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Classe CRigaDocumentoAmministrativo
'Data Creazione 4/11/99
Option Explicit

Public SQL_GETLIST As String
Public SQL_FIND As String
Public SQL_UPDATE As String
Public SQL_INSERT As String
Public SQL_GET_TS As String
Public SQL_GET_SEARCH_LIST As String
Public SQL_GET_SEARCH_LIST_GROUPBY As String
Private SQL_GETLIST_2 As String

Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1
Private mcolListFields As Collection
Private mvarFieldsIndexes As Variant

Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)

Public Function ComputeImponibileRiga(PrezzoUnitario As Variant, _
                                      Quantità As Variant, _
                                      Optional Sconto1 As Variant = 0, _
                                      Optional Sconto2 As Variant = 0, _
                                      Optional Sconto3 As Variant = 0, _
                                      Optional ScontoPerc1 As Variant = 0, _
                                      Optional ScontoPerc2 As Variant = 0, _
                                      Optional ScontoPerc3 As Variant = 0, _
                                      Optional ErrDescr As Variant) As Variant
    On Error GoTo ErrHandler
    Dim strError As String
    Dim varResult As Variant
    
    varResult = Round((PrezzoUnitario - Sconto1 - Sconto2 - Sconto3) * Quantità, 2)
    
    ComputeImponibileRiga = varResult
    Exit Function
ErrHandler:
    ErrDescr = strError
    ComputeImponibileRiga = Empty
End Function


Public Property Get SearchListFields() As Collection
    Set SearchListFields = mcolListFields
End Property
Public Function GetList(Optional ActiveRecordsOnly As Boolean = False) As Boolean
    GetList = DS.GetList(SQL_GETLIST, ActiveRecordsOnly)
End Function
Public Function GetList2(Optional ActiveRecordsOnly As Boolean = False) As Boolean
    GetList2 = DS.GetList(SQL_GETLIST_2, ActiveRecordsOnly)
End Function

Private Sub DS_ClientClear()
    RaiseEvent ClientClear
End Sub
Public Function Find() As Boolean
    Find = DS.Find(Me.SQL_FIND)
End Function
Public Function Insert() As Boolean
'Dim UM As String
'    UM = DS.Field("Unità di Misura").NewValue
'    Select Case UM
'        Case "L"
'            DS.Field("Imponibile").NewValue = (DS.Field("Prezzo Unitario").NewValue * DS.Field("Quantità L").NewValue) - DS.Field("Importo Sconto").NewValue
'        Case "KG"
'            DS.Field("Imponibile").NewValue = (DS.Field("Prezzo Unitario").NewValue * DS.Field("Quantità Kg").NewValue) - DS.Field("Importo Sconto").NewValue
'        Case "CNF"
'            DS.Field("Imponibile").NewValue = (DS.Field("Prezzo Unitario").NewValue * DS.Field("Quantità Cnf").NewValue) - DS.Field("Importo Sconto").NewValue
'        Case "PZ"
'            DS.Field("Imponibile").NewValue = (DS.Field("Prezzo Unitario").NewValue * DS.Field("Quantità Pz").NewValue) - DS.Field("Importo Sconto").NewValue
'    End Select
'
    ' Flag Tipo Riga - default value
    If Not DS.Fields("21").Updated Then
        DS.Fields("21").NewValue = "N"
        DS.Fields("21").Updated = True
    End If

    ' Se non c'è descrizione, imposto la riga come riga "NORMALE"
    If Len(DS.Fields("4").NewValue) = 0 Then
        DS.Fields("21").NewValue = "N"
        DS.Fields("21").Updated = True
    End If
    Insert = DS.Insert(Me.SQL_INSERT)
End Function
Public Sub BeginSave(ByRef blnConcurErr As Boolean)




    DS.BeginSave SQL_UPDATE, SQL_GET_TS, blnConcurErr
End Sub
Private Sub Class_Initialize()
    Dim i As Integer
    SQL_FIND = " SELECT DARIDAEE, DARIDAND, DARIDDAM, DARIDDAR, " & _
               "        DARDESCR, DARQTAKG, DARQTALT, DARQTAPZ, DARQTACF, " & _
               "        DARIDAUM, DARPZUNI, DARIDAIV, DARNRSC1, DARNRSC2, " & _
               "        DARNRSC3, DARNRSV1, DARNRSV2, DARNRSV3, DARIDAPR, " & _
               "        DARIDAIM, DARIDAPD, DARFLTRG, DARNRIMP, DARNRIMS, " & _
               "        DARIDVVT, DARIDVVC, DARIDVVP, DARDENAM, DARDEN15, " & _
               "        DARIMACC, DARIDADD, DARANNUL, TIMESTMP  " & _
               "   FROM docudar0 WHERE "
               
    SQL_GETLIST = " SELECT DARIDAEE, DARIDAND, DARIDDAM, DARIDDAR, DARIDAPR, " & _
                  " APRDESCO, DARQTALT, DARQTAKG, DARQTAPZ, DARQTACF, " & _
                  " DARIDAUM, DARPZUNI, DARIDAIV, DARANNUL " & _
                  " FROM DOCUDAR0 LEFT OUTER JOIN ANAGAPR0 ON DARIDAPR = APRIDAPR "
    SQL_GETLIST_2 = " SELECT DARIDAEE, DARIDAND, DARIDDAM, DARIDDAR, DARIDAPR, " & _
                  " DARDESCR, DARQTALT, DARQTAKG, DARQTAPZ, DARQTACF, " & _
                  " DARIDAUM, DARPZUNI, DARIDAIV, DARANNUL " & _
                  " FROM DOCUDAR0 LEFT OUTER JOIN ANAGAPR0 ON DARIDAPR = APRIDAPR "

    SQL_UPDATE = "UPDATE docudar0 SET "
    
    SQL_INSERT = "INSERT INTO docudar0 "
    
    SQL_GET_TS = "SELECT timestmp FROM docudar0 WHERE "
    
    
    SQL_GET_SEARCH_LIST = "SELECT DARIDAEE,DARIDAND,DARIDDAG,DARIDDAR,DARDESCR,DARANNUL FROM docudar0 "
    
    mvarFieldsIndexes = Array("0", "1", "2", "3", "4", "24")
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i
    
    Set DS = New CDataStruct
    
    With DS.Fields
        
        .Add "Ente Emittente", "docudar0.daridaee", 0, True, 3, "0"
        With .Item("0")
            .ForeignTableEntity = "AEE"
            .ForeignTable = "anagaee0"
            .ForeignTableFld = "aeeidaee"
            .ForeignTableDescField = "aeedescr"
            .ForeignTableDeleteFld = "aeeannul"
        End With
        .Add "Anno Documento", "docudar0.daridand", 0, True, 4, "1"
        .Add "Progressivo Documento", "docudar0.dariddam", 0, True, 7, "2"
        
        .Add "Numero Riga Documento", "docudar0.dariddar", 1, True, 3, "3"
        
        .Add "Descrizione Riga Documento", "docudar0.dardescr", 0, , 64, "4"
        
        .Add "Quantità Kg", "docudar0.darqtakg", 1, , 15, "5"
        .Add "Quantità L", "docudar0.darqtalt", 1, , 15, "6"
        .Add "Quantità Pz", "docudar0.darqtapz", 1, , 15, "7"
        .Add "Quantità Cnf", "docudar0.darqtacf", 1, , 15, "8"
        
        .Add "Unità di Misura", "docudar0.daridaum", 0, , 3, "9"
        With .Item("9")
            .ForeignTableEntity = "AUM"
            .ForeignTable = "anagaum0"
            .ForeignTableFld = "aumidaum"
            .ForeignTableDescField = "aumdescr"
            .ForeignTableDeleteFld = "aumannul"
        End With
        .Add "Prezzo Unitario", "docudar0.darpzuni", 1, , 15, "10"
        
        .Add "Aliquota IVA", "docudar0.daridaiv", 0, , 3, "11"
        With .Item("11")
            .ForeignTableEntity = "AIV"
            .ForeignTable = "anagaiv0"
            .ForeignTableFld = "aividaiv"
            .ForeignTableDescField = "aivdescr"
            .ForeignTableDeleteFld = "aivannul"
        End With
        .Add "Sconto 1", "docudar0.darnrsc1", 1, , 7, "12"
        .Add "Sconto 2", "docudar0.darnrsc2", 1, , 7, "13"
        .Add "Sconto 3", "docudar0.darnrsc3", 1, , 7, "14"
        .Add "Sconto Valore 1", "docudar0.darnrsv1", 1, , 13, "15"
        .Add "Sconto Valore 2", "docudar0.darnrsv2", 1, , 13, "16"
        .Add "Sconto Valore 3", "docudar0.darnrsv3", 1, , 13, "17"
        
        .Add "Prodotto", "docudar0.daridapr", 0, , 6, "18"
        With .Item("18")
            .ForeignTableEntity = "APR"
            .ForeignTable = "anagapr0"
            .ForeignTableFld = "apridapr"
            .ForeignTableDescField = "aprdesco"
            .ForeignTableDeleteFld = "aprannul"
        End With
        .Add "Imballo", "docudar0.daridaim", 0, , 3, "19"
        With .Item("19")
            .ForeignTableEntity = "AIM"
            .ForeignTable = "anagaim0"
            .ForeignTableFld = "aimidaim"
            .ForeignTableDescField = "aimdescr"
            .ForeignTableDeleteFld = "aimannul"
        End With
        .Add "Posizione Doganale", "docudar0.daridapd", 0, , 3, "20"
        With .Item("20")
            .ForeignTableEntity = "APD"
            .ForeignTable = "anagapd0"
            .ForeignTableFld = "apdidapd"
            .ForeignTableDescField = "apddescr"
            .ForeignTableDeleteFld = "apdannul"
        End With
        .Add "Flag Tipo Riga", "docudar0.darfltrg", 0, , 1, "21"
        
        .Add "Imponibile", "docudar0.darnrimp", 1, , 15, "22"
        .Add "Importo Sconto", "docudar0.darnrims", 1, , 15, "23"
        
        .Add "Viaggio", "docudar0.daridvvt", 0, False, 7, "25", , True
        .Add "Consegna", "docudar0.daridvvc", 0, False, 3, "26", , True
        .Add "Scarico", "docudar0.daridvvp", 0, False, 3, "27", , True
        
        ' *** FM 19/09/01
'        .Add "Densità 15", "docudar0.darden15", 1, False, 7, "28"
'        .Add "Densità Ambiente", "docudar0.dardenam", 1, False, 7, "29"

        .Add "Densità Ambiente", "docudar0.dardenam", 1, False, 7, "28"
        .Add "Densità 15", "docudar0.darden15", 1, False, 7, "29"
        ' ***
        .Add "Importo Accisa", "docudar0.darimacc", 1, , 15, "30"
        
        .Add "Addebito", "docudar0.daridadd", 0, , 3, "31"
        With .Item("31")
            .ForeignTableEntity = "ADD"
            .ForeignTable = "anagadd0"
            .ForeignTableFld = "addidadd"
            .ForeignTableDescField = "adddescr"
            .ForeignTableDeleteFld = "addannul"
        End With
        .Add "Annullato", "docudar0.darannul", 3, , 1, "24"
    End With
    With DS.KeyFields    ' CAMPI CHIAVE
        .Add "0", "1"    ' Ente Emittente
        .Add "1", "2"    ' Anno documento
        .Add "2", "3"    ' Progressivo documento
        .Add "3", "4"    ' Numero riga documento
    End With
    DS.CancelField = "docudar0.darannul"
End Sub
Private Sub Class_Terminate()
    Set DS = Nothing
End Sub
Private Sub DS_ClientRefresh()
    RaiseEvent ClientRefresh
End Sub



