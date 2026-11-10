VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CViaggioEvento"
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
Public SQL_GET_SELECTION_EXPORT As String
Public SQL_GET_SELECTION_EXPORT_AGIP As String

Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1
Private mcolListFields As Collection

Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)

Public Function Find() As Boolean
    Find = DS.Find(Me.SQL_FIND)
End Function

Public Function GetList(Optional ActiveRecordsOnly As Boolean = False) As Boolean
    GetList = DS.GetList(SQL_GETLIST, ActiveRecordsOnly)
End Function

Public Function GetSelectionExport(IDViaggio As Variant) As Boolean
    On Error GoTo ErrHandler
    
    With DS
        .Fields.ResetFields
        .Fields("0").FindValue = IDViaggio
        .Fields("38").FindValue = "N"
        GetSelectionExport = .GetSelection(SQL_GET_SELECTION_EXPORT, True)
    End With
    
    Exit Function
ErrHandler:
    GetSelectionExport = False
End Function

Public Function GetSelectionExportAgip(Optional ActiveRecordsOnly As Boolean = False, _
            Optional OrderFields As Variant, _
            Optional MaxRecords As Long = 0) As Boolean
    
On Error GoTo ErrHandler

    Dim varOrderBuffer(0 To 3) As Variant

    If IsEmpty(OrderFields) Then
        varOrderBuffer(0) = "0"
        varOrderBuffer(1) = "1"
        varOrderBuffer(2) = "2"
        varOrderBuffer(3) = "3"
        GetSelectionExportAgip = DS.GetSelection(SQL_GET_SELECTION_EXPORT_AGIP, ActiveRecordsOnly, varOrderBuffer, MaxRecords)
    Else
        GetSelectionExportAgip = DS.GetSelection(SQL_GET_SELECTION_EXPORT_AGIP, ActiveRecordsOnly, OrderFields, MaxRecords)
    End If

    Exit Function
ErrHandler:
    GetSelectionExportAgip = False
End Function

Public Function SetFlagTrasmissione(NewValue As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim strIDViaggio As String
    Dim blnResult As Boolean
    Dim blnConcurErr As Boolean
    
    SetFlagTrasmissione = True
    With DS.Fields
        .Item("38").NewValue = NewValue
        .Item("38").Updated = True
    End With
    blnResult = DS.BeginSave(SQL_UPDATE, SQL_GET_TS, blnConcurErr)
    If blnResult Then
        DS.DA.EndTrans 0
    Else
        DS.DA.EndTrans 1
    End If
    Exit Function
ErrHandler:
    DS.DA.EndTrans 1
    SetFlagTrasmissione = False
End Function

Public Function Insert() As Boolean
    DS.Fields("38").NewValue = "N"
    DS.Fields("38").Updated = True
    Insert = DS.Insert(Me.SQL_INSERT)
End Function


Public Function BeginSave(ByRef blnConcurErr As Boolean) As Boolean
    Dim blnSaveOk As Boolean
    
    DS.Fields("38").NewValue = "N"
    DS.Fields("38").Updated = True
    blnSaveOk = DS.BeginSave(SQL_UPDATE, SQL_GET_TS, blnConcurErr)
    
    BeginSave = True
    
End Function

Private Sub Class_Initialize()
    
    SQL_FIND = "SELECT VPIIDVVT, VPIIDVVC, VPIIDVVP, VPIIDVPI, " & _
               "       VPIIDATI, VPIIDAMI, VPIDTINC, VPIHRINI, " & _
               "       VPIHRFIN, VPIIDORP, VPIIDTIP, VPIIDAPR, " & _
               "       VPIQTAPR, VPIIDAUM, VPIIDPC1, VPIQTPC1, " & _
               "       VPIUMPC1, VPIIDPC2, VPIQTPC2, VPIUMPC2, " & _
               "       VPIIDPC3, VPIQTPC3, VPIUMPC3, VPIIDPC4, " & _
               "       VPIQTPC4, VPIUMPC4, VPIMISSP, VPIIDDER, " & _
               "       VPIDTRIE, VPIIDADD, VPIDTCOD, VPIHRIPA, " & _
               "       VPIHRFPA, VPIDTCOI, VPIFLREI, VPINOTA1, " & _
               "       VPIFLSTR, VPIIDASC, VPIANNUL, timestmp FROM viagvpi0 WHERE "
    
    SQL_GET_SELECTION_EXPORT = " SELECT VPIIDATI, VPIIDAMI, VPIDTINC, VPIHRINI, " & _
                                "       VPIHRFIN, VPIIDORP, VPIIDTIP, VPIIDAPR, " & _
                                "       VPIQTAPR, VPIIDAUM, VPIIDPC1, VPIQTPC1, " & _
                                "       VPIUMPC1, VPIIDPC2, VPIQTPC2, VPIUMPC2, " & _
                                "       VPIIDPC3, VPIQTPC3, VPIUMPC3, VPIIDPC4, " & _
                                "       VPIQTPC4, VPIUMPC4, VPIMISSP, VPIIDDER, " & _
                                "       VPIDTRIE, VPIIDADD, VPIDTCOD, VPIHRIPA, " & _
                                "       VPIHRFPA, VPIDTCOI, VPIFLREI, VPINOTA1, " & _
                                "       OCTRIFE1, OCRRIFE2, OCTIDADZ, VPIIDORP, " & _
                                "       VPIIDTIP, VPIIDVVT, VPIIDVVC, VPIIDVVP, " & _
                                "       VPIIDVPI, OCTRAUT4, OCTIDAEE " & _
                                "  FROM VIAGVVP0 INNER JOIN " & _
                                "       VIAGVPI0 ON VVPIDVVT = VPIIDVVT AND VVPIDVVC = VPIIDVVC AND VVPIDVVP = VPIIDVVP INNER JOIN " & _
                                "       ORDIOCR0 ON VVPIDOCT = OCRIDOCT AND VVPIDOCR = OCRIDOCR INNER JOIN " & _
                                "       ORDIOCT0 ON OCRIDOCT = OCTIDOCT"
    
    SQL_GET_SELECTION_EXPORT_AGIP = "SELECT VIAGVPI0.VPIIDAMI, VIAGVPI0.VPINOTA1, VIAGVVT0.VVTDTINV, VIAGVVT0.VVTIDAUT, ANAGAUT0.AUTBADGE, VIAGVVT0.VVTSEQVI, ORDIOCT0.OCTIDAZO," & _
        " VIAGVVT0.VVTIDADE, ANAGADE0.ADEIDARP, ORDIOCT0.OCTRIFE1, VIAGVPI0.VPIIDAPR, ANAGADZ0.ADZIDADZ, ANAGADZ0.ADZRASOC," & _
        " ORDIOCT0.OCTIDATV , VIAGVPI0.VPIIDVVT, VIAGVPI0.VPIIDVVC, VIAGVPI0.VPIIDVVP, VIAGVPI0.VPIIDVPI, VIAGVVC0.VVCNRSEQ, VPIFLSTR, " & _
        " VVPIDAIM, VVPIDAPD, VVPIDOCT, VVPIDOCR, VVPQCOLT FROM ANAGAUT0 INNER JOIN" & _
        " VIAGVVT0 ON ANAGAUT0.AUTIDAUT = VIAGVVT0.VVTIDAUT INNER JOIN" & _
        " VIAGVVC0 ON VIAGVVT0.VVTIDVVT = VIAGVVC0.VVCIDVVT INNER JOIN" & _
        " VIAGVVP0 ON VIAGVVC0.VVCIDVVT = VIAGVVP0.VVPIDVVT AND VIAGVVC0.VVCIDVVC = VIAGVVP0.VVPIDVVC INNER JOIN" & _
        " VIAGVPI0 ON VIAGVVP0.VVPIDVVT = VIAGVPI0.VPIIDVVT AND VIAGVVP0.VVPIDVVC = VIAGVPI0.VPIIDVVC AND" & _
        " VIAGVVP0.VVPIDVVP = VIAGVPI0.VPIIDVVP INNER JOIN" & _
        " ANAGADE0 ON VIAGVVT0.VVTIDADE = ANAGADE0.ADEIDADE INNER JOIN" & _
        " ANAGADZ0 ON VIAGVVC0.VVCIDADZ = ANAGADZ0.ADZIDADZ INNER JOIN" & _
        " ORDIOCT0 ON VIAGVVC0.VVCIDOCT = ORDIOCT0.OCTIDOCT"

    SQL_GETLIST = "SELECT VPIIDVVT, VPIIDVVC, VPIIDVPI, VPIIDVPI FROM viagvpi0 "
    SQL_UPDATE = "UPDATE viagvpi0 SET "
    SQL_INSERT = "INSERT INTO viagvpi0 "
    SQL_GET_TS = "SELECT timestmp FROM viagvpi0 WHERE "
    
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Numero Viaggio", "viagvpi0.vpiidvvt", 0, True, 7, "0", , True
        .Add "Numero Consegna", "viagvpi0.vpiidvvc", 0, True, 3, "1", , True
        .Add "Numero Scarico", "viagvpi0.vpiidvvp", 0, True, 3, "2", , True
        .Add "Numero Riga", "viagvpi0.vpiidvpi", 0, True, 3, "3", , True
        
        .Add "Tipo evento", "viagvpi0.VPIIDATI", 0, True, 6, "4"
        With .Item("4")
            .ForeignTableEntity = "ATE"
            .ForeignTable = "anagate0"
            .ForeignTableFld = "ateidate"
            .ForeignTableDescField = "atedescr"
            .ForeignTableDeleteFld = "ateannul"
        End With
        .Add "Motivo evento", "viagvpi0.VPIIDAMI", 0, True, 3, "5"
        With .Item("5")
            .ForeignTableEntity = "AME"
            .ForeignTable = "anagame0"
            .ForeignTableFld = "ameidame"
            .ForeignTableDescField = "amedescr"
            .ForeignTableDeleteFld = "ameannul"
        End With
        .Add "Data evento", "viagvpi0.VPIDTINC", 2, True, 10, "6"
        .Add "Ora inizio evento", "viagvpi0.VPIHRINI", 4, False, 5, "7"
        .Add "Ora fine evento", "viagvpi0.VPIHRFIN", 4, False, 5, "8"
        .Add "Ordine precedente", "viagvpi0.VPIIDORP", 0, False, 15, "9"
        .Add "Tipo evento precedente", "viagvpi0.VPIIDTIP", 0, False, 3, "10"
        With .Item("10")
            .ForeignTableEntity = "ATE"
            .ForeignTable = "anagate0"
            .ForeignTableFld = "ateidate"
            .ForeignTableDescField = "atedescr"
            .ForeignTableDeleteFld = "ateannul"
        End With
        .Add "Codice prodotto", "viagvpi0.VPIIDAPR", 0, True, 6, "11", , True
        With .Item("11")
            .ForeignTableEntity = "APR"
            .ForeignTable = "anagapr0"
            .ForeignTableFld = "apridapr"
            .ForeignTableDescField = "aprdesco"
            .ForeignTableDeleteFld = "aprannul"
        End With
        .Add "Quantita Prodotto", "viagvpi0.VPIQTAPR", 1, False, 15, "12"
        .Add "Unità di misura Prodotto", "viagvpi0.VPIIDAUM", 0, False, 3, "13"
        With .Item("13")
            .ForeignTableEntity = "AUM"
            .ForeignTable = "anagaum0"
            .ForeignTableFld = "aumidaum"
            .ForeignTableDescField = "aumdescr"
            .ForeignTableDeleteFld = "aumannul"
        End With
        
        .Add "Codice prodotto contaminato 1", "viagvpi0.VPIIDPC1", 0, False, 6, "14"
        With .Item("14")
            .ForeignTableEntity = "APR"
            .ForeignTable = "anagapr0"
            .ForeignTableFld = "apridapr"
            .ForeignTableDescField = "aprdesco"
            .ForeignTableDeleteFld = "aprannul"
        End With
        .Add "Quantita Prodotto contaminato 1", "viagvpi0.VPIQTPC1", 1, False, 15, "15"
        .Add "Unità di misura Prodotto contaminato 1", "viagvpi0.VPIUMPC1", 0, False, 3, "16"
        With .Item("16")
            .ForeignTableEntity = "AUM"
            .ForeignTable = "anagaum0"
            .ForeignTableFld = "aumidaum"
            .ForeignTableDescField = "aumdescr"
            .ForeignTableDeleteFld = "aumannul"
        End With
        .Add "Codice prodotto contaminato 2", "viagvpi0.VPIIDPC2", 0, False, 6, "17"
        With .Item("17")
            .ForeignTableEntity = "APR"
            .ForeignTable = "anagapr0"
            .ForeignTableFld = "apridapr"
            .ForeignTableDescField = "aprdesco"
            .ForeignTableDeleteFld = "aprannul"
        End With
        .Add "Quantita Prodotto contaminato 2", "viagvpi0.VPIQTPC2", 1, False, 15, "18"
        .Add "Unità di misura Prodotto contaminato 2", "viagvpi0.VPIUMPC2", 0, False, 3, "19"
        With .Item("19")
            .ForeignTableEntity = "AUM"
            .ForeignTable = "anagaum0"
            .ForeignTableFld = "aumidaum"
            .ForeignTableDescField = "aumdescr"
            .ForeignTableDeleteFld = "aumannul"
        End With
        .Add "Codice prodotto contaminato 3", "viagvpi0.VPIIDPC3", 0, False, 6, "20"
        With .Item("20")
            .ForeignTableEntity = "APR"
            .ForeignTable = "anagapr0"
            .ForeignTableFld = "apridapr"
            .ForeignTableDescField = "aprdesco"
            .ForeignTableDeleteFld = "aprannul"
        End With
        .Add "Quantita Prodotto contaminato 3", "viagvpi0.VPIQTPC3", 1, False, 15, "21"
        .Add "Unità di misura Prodotto contaminato 3", "viagvpi0.VPIUMPC3", 0, False, 3, "22"
        With .Item("22")
            .ForeignTableEntity = "AUM"
            .ForeignTable = "anagaum0"
            .ForeignTableFld = "aumidaum"
            .ForeignTableDescField = "aumdescr"
            .ForeignTableDeleteFld = "aumannul"
        End With
        .Add "Codice prodotto contaminato 4", "viagvpi0.VPIIDPC4", 0, False, 6, "23"
        With .Item("23")
            .ForeignTableEntity = "APR"
            .ForeignTable = "anagapr0"
            .ForeignTableFld = "apridapr"
            .ForeignTableDescField = "aprdesco"
            .ForeignTableDeleteFld = "aprannul"
        End With
        .Add "Quantita Prodotto contaminato 4", "viagvpi0.VPIQTPC4", 1, False, 15, "24"
        .Add "Unità di misura Prodotto contaminato 4", "viagvpi0.VPIUMPC4", 0, False, 3, "25"
        With .Item("25")
            .ForeignTableEntity = "AUM"
            .ForeignTable = "anagaum0"
            .ForeignTableFld = "aumidaum"
            .ForeignTableDescField = "aumdescr"
            .ForeignTableDeleteFld = "aumannul"
        End With
        
        .Add "Misure adottate per spandimento", "viagvpi0.VPIMISSP", 0, False, 50, "26"
    
        .Add "Deposito rientro", "viagvpi0.VPIIDDER", 0, False, 3, "27"
        With .Item("27")
            .ForeignTableEntity = "ADE"
            .ForeignTable = "anagade0"
            .ForeignTableFld = "adeidade"
            .ForeignTableDescField = "adedescr"
            .ForeignTableDeleteFld = "adeannul"
        End With
        .Add "Data rientro", "viagvpi0.VPIDTRIE", 2, False, 10, "28"
        
        .Add "Destinazione dirottamento", "viagvpi0.VPIIDADD", 0, False, 10, "29"
        With .Item("29")
            .ForeignTableEntity = "ADZ"
            .ForeignTable = "anagadz0"
            .ForeignTableFld = "adzidadz"
            .ForeignTableDescField = "adzrasoc"
            .ForeignTableDeleteFld = "adzannul"
        End With
        .Add "Data dirottamento", "viagvpi0.VPIDTCOD", 2, False, 10, "30"
       
        .Add "Ora inizio paralisi", "viagvpi0.VPIHRIPA", 4, False, 5, "31"
        .Add "Ora fine paralisi", "viagvpi0.VPIHRFPA", 4, False, 5, "32"
        .Add "Data comunicazione", "viagvpi0.VPIDTCOI", 2, True, 10, "35"
        .Add "Responsabile comunicazione", "viagvpi0.VPIFLREI", 0, True, 2, "36"
        .Add "Osservazioni", "viagvpi0.VPINOTA1", 0, False, 200, "37"
        .Add "Stato trasmissione", "viagvpi0.VPIFLSTR", 0, True, 2, "38"
        .Add "Società", "viagvpi0.VPIIDASC", 0, True, 3, "40"
        With .Item("40")
            .ForeignTableEntity = "ASC"
            .ForeignTable = "anagasc0"
            .ForeignTableFld = "ascidasc"
            .ForeignTableDescField = "ascrasoc"
            .ForeignTableDeleteFld = "ascannul"
        End With

        .Add "Annullato", "viagvpi0.vpiannul", 3, , , "39"
    End With
    With DS.KeyFields
        .Add "0", "1"
        .Add "1", "2"
        .Add "2", "3"
        .Add "3", "4"
        .Add "38", "5"
        .Add "39", "6"
        .Add "40", "7"
  End With
    DS.CancelField = "viagvpi0.vpiannul"
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

