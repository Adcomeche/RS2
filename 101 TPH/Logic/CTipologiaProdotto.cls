VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTipologiaProdotto"
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

Private mcolListFields As Collection
Private mvarFieldsIndexes As Variant
Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1

Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)


' (DB) 11/08/2005
Public Sub TrovaTipologiaProdDaRaggrProd(ByVal pRaggrProdotto As String, ByRef rTipologiaProdotto As Variant)
    Const GET_TTP_FROM_ARR = _
        "SELECT TTPIDTTP, APRIDAPR " & _
        "  FROM ANAGAPR0 INNER JOIN ANAGTTP0 ON APRIDTTP = TTPIDTTP " & _
        " WHERE APRIDARR = ? "
    Dim varBuf As Variant
    Dim lngHowMany As Long
    On Error GoTo ErrHandler
    
    rTipologiaProdotto = Empty
    lngHowMany = 0
    If Not DS.DA.QueryCommand(GET_TTP_FROM_ARR, varBuf, lngHowMany, pRaggrProdotto) Then
        GoTo ErrHandler
    End If
    If lngHowMany > 0 Then
        rTipologiaProdotto = varBuf(0, 0)
    End If
Exit Sub
ErrHandler:
    On Error Resume Next
    
    rTipologiaProdotto = Empty
End Sub

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
    Dim i As Integer
    
    SQL_FIND = "SELECT ttpidttp, ttpdescr, ttpflges, ttpidaum, " & _
               "       ttpidums, ttpannul, timestmp FROM anagttp0 WHERE "
    SQL_GETLIST = "SELECT ttpidttp, ttpdescr FROM anagttp0 "
    SQL_UPDATE = "UPDATE anagttp0 SET "
    SQL_INSERT = "INSERT INTO anagttp0 "
    SQL_GET_TS = "SELECT timestmp FROM anagttp0 WHERE "
    SQL_GET_SEARCH_LIST = "SELECT ttpidttp, ttpdescr, ttpannul FROM anagttp0 "
    
    mvarFieldsIndexes = Array("Codice", "Descrizione", "Annullato")
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i
    
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Codice", "anagttp0.ttpidttp", 0, True, 3, "Codice"
        .Add "Descrizione", "anagttp0.ttpdescr", 0, True, 25, "Descrizione"
        .Add "Flag Gestione", "anagttp0.ttpflges", 0, False, 2, "Gestione"
        .Add "Unità di misura", "anagttp0.ttpidaum", 0, False, 3, "UM"
        With .Item("UM")
            .ForeignTableEntity = "AUM"
            .ForeignTable = "anagaum0"
            .ForeignTableFld = "aumidaum"
            .ForeignTableDescField = "aumdescr"
            .ForeignTableDeleteFld = "aumannul"
        End With
        .Add "Unità di misura stoccaggio", "anagttp0.ttpidums", 0, False, 3, "UM Stoccaggio"
        With .Item("UM Stoccaggio")
            .ForeignTableEntity = "AUM"
            .ForeignTable = "anagaum0"
            .ForeignTableFld = "aumidaum"
            .ForeignTableDescField = "aumdescr"
            .ForeignTableDeleteFld = "aumannul"
        End With
        .Add "Annullato", "anagttp0.ttpannul", 3, , , "Annullato"
    End With
    With DS.KeyFields
        .Add "Codice", "1"
    End With
    DS.CancelField = "anagttp0.ttpannul"
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


Public Property Get SearchListFields() As Collection
    Set SearchListFields = mcolListFields
End Property


