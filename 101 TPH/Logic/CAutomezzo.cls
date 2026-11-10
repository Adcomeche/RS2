VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAutomezzo"
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

Private SQL_FIND_CONSUMI As String
Private SQL_INSERT_CONSUMI As String
Private SQL_UPDATE_CONSUMI As String

Private mcolListFields As Collection
Private mvarFieldsIndexes As Variant
Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1

Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)

Private Const SQL_FIND_BY_EXT_REF = "SELECT AAMIDAAM From ANAGAAM0 WHERE (REPLACE(REPLACE(AAMIDAAM, ' ', ''), '-', '') = ?)"

' -- Invoca la funzione di controllo sulla possibilità di utilizzo in una data

Public Function CheckScadenze(varIDKEY As Variant, dtDateToCheck As Date, varIDASC As Variant, ByRef varResult As Variant) As Boolean

    On Error GoTo ErrHandler

 
    Dim SCASelect  As CScadenzeSelezione
    Set SCASelect = New CScadenzeSelezione

    With SCASelect.DS
        Set .DA = DS.DA
        Set .ErrH = DS.ErrH
    End With
    CheckScadenze = SCASelect.blnCheckScadenze(varIDKEY, dtDateToCheck, "AM", varResult, varIDASC)
    Set SCASelect = Nothing

    Exit Function

 

ErrHandler:

    Set SCASelect = Nothing

    CheckScadenze = False

End Function



Public Function InsertConsumi(IDAutomezzo As Variant, Data As Variant, ConsumoInterno As Variant, ConsumoEsterno As Variant) As Boolean
    On Error GoTo ErrHandler
    
    Dim varBuffer As Variant
    Dim lngHowMany As Long
    
    If Not IsDate(Data) Then
        GoTo ErrHandler
    End If
    If Not IsNumeric(ConsumoInterno) Then
        GoTo ErrHandler
    End If
    If Not IsNumeric(ConsumoEsterno) Then
        GoTo ErrHandler
    End If
    
    DS.DA.BeginTransaction
    
'    DS.Field("Codice").FindValue = IDAutomezzo
'    If Not Me.Find Then
'        GoTo ErrHandler
'    End If
'    If DS.HowMany = 0 Then
'        GoTo ErrHandler
'    End If
    
    If Not DS.DA.QueryCommand(SQL_FIND_CONSUMI & DS.DA.MakeSQL(Data, dtDate), varBuffer, lngHowMany, IDAutomezzo) Then
        GoTo ErrHandler
    End If
    
    If varBuffer(0, 0) = 0 Then
        If Not DS.DA.ExecuteCommand(SQL_INSERT_CONSUMI & DS.DA.MakeSQL(Data, dtDate) & ", " & DS.DA.MakeSQL(ConsumoInterno, dtNum) & ", " & DS.DA.MakeSQL(ConsumoEsterno, dtNum) & ")", IDAutomezzo) Then
            GoTo ErrHandler
        End If
    Else
        If Not DS.DA.ExecuteCommand(SQL_UPDATE_CONSUMI, ConsumoInterno, ConsumoEsterno, IDAutomezzo, Data) Then
            GoTo ErrHandler
        End If
    End If
    
    DS.DA.EndTrans etCommit
    InsertConsumi = True
    Exit Function
ErrHandler:
    DS.DA.EndTrans etRollback
    InsertConsumi = False
End Function

Public Function GetIDByExtRef(Ref As String) As String
    On Error GoTo ErrHandler
    Dim blnResult As Boolean
    Dim varBuffer As Variant
    Dim lngHowMany As Long
    
    Ref = Replace(Replace(Ref, "-", ""), " ", "")
    blnResult = DS.DA.QueryCommand(SQL_FIND_BY_EXT_REF, varBuffer, lngHowMany, Ref)
    If lngHowMany = 0 Then
        GoTo ErrHandler
    End If
    GetIDByExtRef = varBuffer(0, 0)
    Exit Function
ErrHandler:
    GetIDByExtRef = vbNullString
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


Public Function BeginSave(ByRef blnConcurErr As Boolean) As Boolean

    ' [09-05-2005 (PM)] Aggiunto il default del satellitare installato
    If (DS.Fields("34").Updated And DS.Fields("34").NewValue = "") Or _
        (Not DS.Fields("34").Updated And DS.Fields("34").Value = "") Then
        DS.Fields("34").NewValue = "N"
        DS.Fields("34").Updated = True
    End If

    BeginSave = DS.BeginSave(SQL_UPDATE, SQL_GET_TS, blnConcurErr)
End Function
Public Function Validate() As Boolean
    On Error GoTo ErrHandler
    Dim blnResult As Boolean
    Dim strMotRim As String
    Dim strErrorMsg As String
    Dim blnValidateOk As Boolean
    
    blnValidateOk = True
    blnResult = DS.Fields.Validate
    If DS.Fields("1").Updated Then
        strMotRim = DS.Fields("1").NewValue
    Else
        strMotRim = DS.Fields("1").Value
    End If
    
    If strMotRim = "R" Then
       If DS.Fields("25").Updated Then
            If DS.Fields("25").NewValue <= 0 Then
                blnValidateOk = False
                strErrorMsg = "La portata dell'automezzo deve essere maggiore di 0"
                If Not DS.Fields.AddValidationError(strErrorMsg) Then
                    GoTo ErrHandler
                End If
            End If
       End If
    End If
    
    Validate = blnValidateOk
    Exit Function
ErrHandler:
    Validate = False
End Function

Public Function FindIDVehiOes(Matricula) As String
    
    Dim rowbuff As Variant
    Dim lng As Long
    
    SQL_FIND = "SELECT AAMIDDEU From ANAGAAM0 WHERE AAMIDAAM = ? "
        If Not DS.DA.QueryCommand(SQL_FIND, rowbuff, lng, Matricula) Then
        End If
    FindIDVehiOes = rowbuff(0, 0)

End Function


Private Sub Class_Initialize()
    Dim i As Integer
    
    SQL_FIND = "SELECT aamidaam, aamfltip, aamdescr, aamidatm, aamfldis," & _
               "aamcap01, aamnum01, aamcap02, aamnum02, aamcap03, aamnum03, " & _
               "aamcap04, aamnum04, aamcap05, aamnum05, aamcap06, aamnum06, " & _
               "aamcap07, aamnum07, aamcap08, aamnum08, aamcap09, aamnum09, " & _
               "aamcap10, aamnum10, aamporkg, aamidats, aamtarap, aamtaral, " & _
               "aamporri, aamcubco, aamcaric, aamflsca, aamflsat, aamflnaz, " & _
               "aamiddeu, aamannul, timestmp FROM anagaam0 WHERE "
    SQL_GETLIST = "SELECT aamidaam, aamdescr FROM anagaam0 "
    SQL_UPDATE = "UPDATE anagaam0 SET "
    SQL_INSERT = "INSERT INTO anagaam0 "
    SQL_GET_TS = "SELECT timestmp FROM anagaam0 WHERE "
    SQL_GET_SEARCH_LIST = "SELECT aamidaam, aamfltip, aamdescr, aamidatm, aamfldis,aamporkg, aamannul FROM anagaam0 "
    
    SQL_FIND_CONSUMI = "SELECT COUNT(*) FROM STIPCAM0 WHERE CAMIDAAM = ? AND CAMDTCAM = "
    SQL_INSERT_CONSUMI = "INSERT INTO STIPCAM0 (CAMIDAAM, CAMDTCAM, CAMNRCIN, CAMNRCES) VALUES (?, "
    SQL_UPDATE_CONSUMI = "UPDATE STIPCAM0 SET CAMNRCIN = ?, CAMNRCES = ? WHERE CAMIDAAM = ? AND CAMDTCAM = ?"
    
    mvarFieldsIndexes = Array("0", "1", "2", "3", "4", "25", "26")
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i
    
    
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Targa", "anagaam0.aamidaam", 0, True, 10, "0"
        .Add "Motrice/Rimorchio", "anagaam0.aamfltip", 0, True, 1, "1"
        .Add "Descrizione", "anagaam0.aamdescr", 0, False, 30, "2"
        .Add "Tipo Automezzo", "anagaam0.aamidatm", 0, True, 3, "3"
        With .Item("3")
            .ForeignTableEntity = "ATM"
            .ForeignTable = "anagatm0"
            .ForeignTableFld = "atmidatm"
            .ForeignTableDescField = "atmdescr"
            .ForeignTableDeleteFld = "atmannul"
        End With
        .Add "Disponibilità", "anagaam0.aamfldis", 0, True, 1, "4"
        For i = 1 To 10
            .Add "Capacità Scomparto " & i, "anagaam0.aamcap" & Format$(i, "00"), 1, False, 5, CStr((i * 2 - 1) + 4)
            .Add "Numero Scomparti " & i, "anagaam0.aamnum" & Format$(i, "00"), 1, False, 2, CStr(i * 2 + 4)
        Next i
        .Add "Portata", "anagaam0.aamporkg", 1, True, 6, "25"
        .Add "Tipo Cisterna", "anagaam0.aamidats", 0, , 3, "27"
        With .Item("27")
            .ForeignTableEntity = "ATS"
            .ForeignTable = "anagats0"
            .ForeignTableFld = "atsidats"
            .ForeignTableDescField = "atsdescr"
            .ForeignTableDeleteFld = "atsannul"
        End With
        .Add "Tara Pesata", "anagaam0.aamtarap", 1, False, 6, "28"
        .Add "Tara Libretto", "anagaam0.aamtaral", 1, False, 6, "29"
        .Add "Portata rimorchiabile", "anagaam0.aamporri", 1, False, 6, "30"
        .Add "Cubatura costruttore", "anagaam0.aamcubco", 1, False, 6, "31"
        .Add "Tipo Carica", "anagaam0.aamcaric", 0, True, 1, "32"
        .Add "Tipo Scarico", "anagaam0.aamflsca", 0, False, 1, "33"
        ' [28-04-05 (PM)] Aggiunto flag presenza terminale satellitare
        .Add "Satellitare installato", "anagaam0.aamflsat", dtAlfa, True, , "34"
        .Add "Flag Nazionalità", "anagaam0.aamflnaz", 0, , 2, "35"
        .Add "Codice satellitare", "anagaam0.aamiddeu", 0, False, 10, "36"
        .Add "Annullato", "anagaam0.aamannul", 3, , , "26"
        
        
    End With
    With DS.KeyFields
        .Add "0", "1"
    End With
    DS.CancelField = "anagaam0.aamannul"
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




