VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDeposito"
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
Public Function DecodeFirmCode(FirmID As String, FirmCode As String, Optional ReturnID As Boolean = False) As String
    On Error GoTo ErrHandler
    
    DecodeFirmCode = DS.DecodeFirmCode(FirmID, FirmCode, Me.SQL_FIND, ReturnID)
    
    Exit Function
ErrHandler:
    DecodeFirmCode = vbNullString
End Function

Public Function DecodeFirmCode2(FirmID As String, FirmCode As String, Optional ReturnID As Boolean = False) As String
    On Error GoTo ErrHandler
    
    Dim strQ As String
    Dim strID As String
    Dim varResult As Variant
    Dim lngHowMany As Long
    Dim i As Long
    
    With DS.MultiSoc.DS
        strQ = DS.MultiSoc.SQL_FIND
        strQ = strQ & .Fields("Società").FieldName & " = " & DS.DA.MakeSQL(FirmID, 0)
        strQ = strQ & " AND " & .Fields("CodPerSoc").FieldName & " = " & DS.DA.MakeSQL(FirmCode, 0)
        DS.DA.QueryData strQ, varResult, lngHowMany
        If lngHowMany > 0 Then
            For i = 1 To lngHowMany
                If varResult(2, i - 1) = FirmCode Then
                    strID = varResult(0, i - 1)
                End If
            Next i
            If Not ReturnID Then
                DS.Fields(DS.KeyFields(1)).FindValue = strID
                If Find() Then
                    DecodeFirmCode2 = DS.Fields("1").Value
                End If
            Else
                DecodeFirmCode2 = strID
            End If
        End If
    End With
    
    Exit Function
ErrHandler:
    DecodeFirmCode2 = vbNullString
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

    If IsNull(DS.Field("Trasmissione KM CEPSA").NewValue) Or _
        IsEmpty(DS.Field("Trasmissione KM CEPSA").NewValue) Or _
        Len(DS.Field("Trasmissione KM CEPSA").NewValue) = 0 Then
        DS.Field("Trasmissione KM CEPSA").NewValue = "D"
        DS.Field("Trasmissione KM CEPSA").Updated = True
    End If
    Insert = DS.Insert(Me.SQL_INSERT)
End Function


Public Sub BeginSave(ByRef blnConcurErr As Boolean)
    DS.BeginSave SQL_UPDATE, SQL_GET_TS, blnConcurErr
End Sub
Private Sub Class_Initialize()
    Dim i As Integer
    
    SQL_FIND = "SELECT adeidade, adedescr, adeindir, adelocal, " & _
               "       adeprovi, adenrcap, adeidasc, adeidazo, " & _
               "       adeidarp, adecodma, adefltkc, adecgpsx, " & _
               "       adecgpsy, adeannul, timestmp FROM anagade0 WHERE "
               
    SQL_GETLIST = "SELECT adeidade, adedescr FROM anagade0 "
    SQL_UPDATE = "UPDATE anagade0 SET "
    SQL_INSERT = "INSERT INTO anagade0 "
    SQL_GET_TS = "SELECT timestmp FROM anagade0 WHERE "
    SQL_GET_SEARCH_LIST = "SELECT adeidade, adedescr, adeidasc, adeidazo, adeindir, adelocal, " & _
               "adeprovi, adenrcap, adeannul FROM anagade0 "
    
    mvarFieldsIndexes = Array("0", "1", "6", "7", "2", "3", "4", "5", "9")
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i
    
    Set DS = New CDataStruct
    Set DS.MultiSoc = New CDepositoSoc
    
    With DS.Fields
        .Add "Codice", "anagade0.adeidade", 0, True, 3, "0"
        .Add "Descrizione", "anagade0.adedescr", 0, True, 25, "1"
        .Add "Indirizzo", "anagade0.adeindir", 0, False, 40, "2"
        .Add "Località", "anagade0.adelocal", 0, False, 25, "3"
        .Add "Provincia", "anagade0.adeprovi", 0, False, 2, "4"
        .Add "CAP", "anagade0.adenrcap", 0, False, 10, "5"
        .Add "Società", "anagade0.adeidasc", 0, True, 3, "6"
        With .Item("6")
            .ForeignTableEntity = "ASC"
            .ForeignTable = "anagasc0"
            .ForeignTableFld = "ascidasc"
            .ForeignTableDescField = "ascrasoc"
            .ForeignTableDeleteFld = "ascannul"
        End With
        .Add "Zona", "anagade0.adeidazo", 0, True, 3, "7"
        With .Item("7")
            .ForeignTableEntity = "AZO"
            .ForeignTable = "anagazo0"
            .ForeignTableFld = "azoidazo"
            .ForeignTableDescField = "azodescr"
            .ForeignTableDeleteFld = "azoannul"
        End With
        .Add "Raggruppamento", "anagade0.adeidarp", 0, False, 3, "8"
        With .Item("8")
            .ForeignTableEntity = "ARP"
            .ForeignTable = "anagarp0"
            .ForeignTableFld = "arpidarp"
            .ForeignTableDescField = "arpdescr"
            .ForeignTableDeleteFld = "arpannul"
        End With
        .Add "Presidio", "anagade0.adecodma", 0, , 15, "10"
        .Add "Trasmissione KM CEPSA", "anagade0.adefltkc", 0, , 2, "11"
        
        ' -- DP 10/07/2009
        ' -- Aggiunti campi coordinate geo per invio su GES
        .Add "Coordinate GPS X", "anagade0.adecgpsx", 0, False, 15, "12"
        .Add "Coordinate GPS Y", "anagade0.adecgpsy", 0, False, 15, "13"
        
        .Add "Annullato", "anagade0.adeannul", 3, , , "9"
    End With
    With DS.KeyFields
        .Add "0", "1"
    End With
    DS.CancelField = "anagade0.adeannul"
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




