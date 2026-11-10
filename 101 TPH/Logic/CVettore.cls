VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVettore"
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
    BeginSave = DS.BeginSave(SQL_UPDATE, SQL_GET_TS, blnConcurErr)
End Function
Private Sub Class_Initialize()
    Dim i As Integer
    
    SQL_FIND = "SELECT aveidave, averasoc, averasag, averasab, " & _
               "aveindir, avelocal, aveprovi, avenrcap, " & _
               "avecofis, avepaiva, aveidamp, aveidaiv, " & _
               "avenrtel, avenota1, avenota2, aveannul, timestmp FROM anagave0 WHERE "
    SQL_GETLIST = "SELECT aclidamp, aclrasoc FROM anagave0 "
    SQL_UPDATE = "UPDATE anagave0 SET "
    SQL_INSERT = "INSERT INTO anagave0 "
    SQL_GET_TS = "SELECT timestmp FROM anagave0 WHERE "
    SQL_GET_SEARCH_LIST = "SELECT aveidave, averasoc, aveindir, avelocal, aveprovi, " & _
                          "avenrcap, aveannul FROM anagave0 "
    Set DS = New CDataStruct
    Set DS.MultiSoc = New CVettoreSoc
    mvarFieldsIndexes = Array("0", "1", "4", "5", "6", "7", "15")
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i
    With DS.Fields
        .Add "Codice", "anagave0.aveidave", 0, True, 3, "0"
        .Add "Ragione Sociale", "anagave0.averasoc", 0, True, 40, "1"
        .Add "Ragione Sociale Aggiuntiva", "anagave0.averasag", 0, False, 40, "2"
        .Add "Ragione Sociale Abbreviata", "anagave0.averasab", 0, False, 6, "3"
        .Add "Indirizzo", "anagave0.aveindir", 0, False, 40, "4"
        .Add "Località", "anagave0.avelocal", 0, False, 25, "5"
        .Add "Provincia", "anagave0.aveprovi", 0, False, 2, "6"
        .Add "CAP", "anagave0.avenrcap", 0, False, 10, "7"
        .Add "Codice Fiscale", "anagave0.avecofis", 0, False, 16, "8"
        .Add "Partita IVA", "anagave0.avepaiva", 0, True, 20, "9"
        .Add "Modalità di Pagamento", "anagave0.aveidamp", 0, True, 3, "10"
        With .Item("10")
            .ForeignTableEntity = "AMP"
            .ForeignTable = "anagamp0"
            .ForeignTableFld = "ampidamp"
            .ForeignTableDescField = "ampdescr"
            .ForeignTableDeleteFld = "ampannul"
        End With
        .Add "IVA", "anagave0.aveidaiv", 0, True, 3, "11"
        With .Item("11")
            .ForeignTableEntity = "AIV"
            .ForeignTable = "anagaiv0"
            .ForeignTableFld = "aividaiv"
            .ForeignTableDescField = "aivdescr"
            .ForeignTableDeleteFld = "aivannul"
        End With
        .Add "Telefono", "anagave0.avenrtel", 0, False, 20, "12"
        .Add "Nota 1", "anagave0.avenota1", 0, False, 40, "13"
        .Add "Nota 2", "anagave0.avenota2", 0, False, 40, "14"
        .Add "Annullato", "anagave0.aveannul", 3, , , "15"
    End With
    With DS.KeyFields
        .Add "0", "1"
    End With
    DS.CancelField = "anagave0.aveannul"
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

Public Function GetNewKey(ByRef strKey As String) As Boolean
On Error GoTo ErrHandler

    Dim AVE As CVettore
    Dim varCounter As Variant
    Dim lngKeyPrefix As Long

    Set AVE = New CVettore
    With AVE.DS
        Set .DA = DS.DA
        Set .ErrH = DS.DA.ErrH
    End With

    Do
        varCounter = DS.DA.GetCounter("anaganu0", "anuprogr", "anuidanu", "AVE", 0)
        If varCounter = -1 Then
            GoTo ErrHandler
        End If
    
        lngKeyPrefix = CLng(Mid(varCounter, Len(varCounter) - 3, 2))
        If lngKeyPrefix <= 9 Then
            strKey = Chr(48 + lngKeyPrefix) & Right(varCounter, 2)
        Else
            strKey = Chr(65 + (lngKeyPrefix - 9)) & Right(varCounter, 2)
        End If
        
        With AVE
            .DS.Fields.ResetFields
            .DS.Fields("0").FindValue = strKey
            If Not .Find Then
                GoTo ErrHandler
            End If

            If .DS.HowMany = 0 Then
                Exit Do
            End If
        End With
    Loop

    Set AVE = Nothing
    GetNewKey = True
    Exit Function

ErrHandler:
    Set AVE = Nothing
    GetNewKey = False
End Function


