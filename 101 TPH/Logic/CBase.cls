VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBase"
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

Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)

Private mcolListFields As Collection
Private mvarFieldsIndexes As Variant

Private SQL_INSERT_DISTANZE As String
Private SQL_UPDATE_DISTANZE As String
Private SQL_DELETE_DISTANZA As String
Private SQL_GET_DISTANZA As String
Private SQL_GET_LIST_DISTANZE As String

Public Function DeleteDistanze(IDBase As String, IDDeposito As String) As Boolean
    On Error GoTo ErrHandler
    
    DeleteDistanze = DS.DA.ExecuteCommand(SQL_DELETE_DISTANZA, IDBase, IDDeposito)
    
    Exit Function
ErrHandler:
    DeleteDistanze = False
End Function

Public Function InsertDistanze(IDBase As String, IDDeposito As String, Minuti As Long) As Boolean
    On Error GoTo ErrHandler
    
    If Not IsNumeric(Minuti) Then Minuti = 0
    
    InsertDistanze = DS.DA.ExecuteCommand(SQL_INSERT_DISTANZE, IDBase, IDDeposito, Minuti)
    
    Exit Function
ErrHandler:
    InsertDistanze = False
End Function

Public Function GetListDistanze(IDDestinazione As String) As Variant
    On Error GoTo ErrHandler
    Dim varBuffer As Variant
    Dim blnResult As Boolean
    
    blnResult = DS.DA.QueryCommand(SQL_GET_LIST_DISTANZE, varBuffer, _
                                   0, IDDestinazione)
    If Not blnResult Then
        GoTo ErrHandler
    End If
    GetListDistanze = varBuffer
    
    Exit Function
ErrHandler:
    GetListDistanze = Empty
End Function

Public Function UpdateDistanze(IDBase As String, IDDeposito As String, Minuti As Long) As Boolean
    On Error GoTo ErrHandler
    
    If Not IsNumeric(Minuti) Then Minuti = 0
    
    UpdateDistanze = DS.DA.ExecuteCommand(SQL_UPDATE_DISTANZE, Minuti, IDBase, IDDeposito)
    
    Exit Function
ErrHandler:
    UpdateDistanze = False
End Function


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
    
    SQL_FIND = "SELECT absidabs, absdescr, absannul, timestmp FROM anagabs0 WHERE "
    SQL_GETLIST = "SELECT absidabs, absdescr FROM anagabs0 "
    SQL_UPDATE = "UPDATE anagabs0 SET "
    SQL_INSERT = "INSERT INTO anagabs0 "
    SQL_GET_TS = "SELECT timestmp FROM anagabs0 WHERE "
    SQL_GET_SEARCH_LIST = "SELECT absidabs, absdescr, absannul FROM anagabs0 "
    
    SQL_GET_LIST_DISTANZE = "SELECT BXDIDADE, BXDNRMIN " & _
                            "  FROM ANAGBXD0 " & _
                            " WHERE BXDIDABS = ? "
    SQL_INSERT_DISTANZE = "INSERT INTO ANAGBXD0 (BXDIDABS, BXDIDADE, BXDNRMIN)" & _
                          " VALUES (?, ?, ?) "
    SQL_GET_DISTANZA = "SELECT BXDNRMIN " & _
                       "  FROM ANAGBXD0 " & _
                       " WHERE BXDIDABS = ? AND BXDIDADE = ? "
    SQL_UPDATE_DISTANZE = "UPDATE ANAGBXD0 " & _
                          "   SET BXDNRMIN = ?  " & _
                          " WHERE BXDIDABS = ? AND BXDIDADE = ? "
    SQL_DELETE_DISTANZA = "DELETE ANAGBXD0 WHERE BXDIDABS = ? AND BXDIDADE = ?"
    
    mvarFieldsIndexes = Array("Codice", "Descrizione", "Annullato")
    Set mcolListFields = New Collection
    For i = 0 To UBound(mvarFieldsIndexes)
        mcolListFields.Add mvarFieldsIndexes(i)
    Next i
    
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Codice", "anagabs0.absidabs", 0, True, 3, "Codice"
        .Add "Descrizione", "anagabs0.absdescr", 0, True, 30, "Descrizione"
        .Add "Annullato", "anagabs0.absannul", 3, , , "Annullato"
    End With
    With DS.KeyFields
        .Add "Codice", "1"
    End With
    DS.CancelField = "anagabs0.absannul"
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


