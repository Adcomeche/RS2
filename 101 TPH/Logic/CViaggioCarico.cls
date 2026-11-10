VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CViaggioCarico"
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
Public SQL_GET_LOAD_DATA As String
Public SQL_INSERT_LOAD_DATA As String
Public SQL_UPDATE_LOAD_DATA As String
Public SQL_DELETE_LOAD_DATA As String

Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1
Private mcolListFields As Collection

Event ClientRefresh()
Event ClientClear()
Event Command(Cmd As FormCommand)

Public Function DeleteLoadData(IDViaggio As Variant) As Boolean
    On Error GoTo ErrHandler
    
'    Dim VVT As CViaggio
'
'    Set VVT = New CViaggio
'    With VVT.DS
'        Set .DA = DS.DA
'        Set .ErrH = DS.ErrH
'    End With
    
    If Not DS.DA.ExecuteCommand(SQL_DELETE_LOAD_DATA, IDViaggio) Then
        GoTo ErrHandler
    End If
    
'    VVT.DS.Fields("0").FindValue = DS.Fields("0").Value
'    VVT.Find
'    VVT.SetFlagTrasmissioneESys "N"
'
    DeleteLoadData = True
    
    Exit Function
    
ErrHandler:
    DeleteLoadData = False
End Function


Public Function GetLoadData(IDViaggio As Variant, _
                            Optional InsertData As Boolean = False) As Variant
    On Error GoTo ErrHandler
    Dim varBuffer As Variant
    Dim lngHowMany As Long
    
    ' --- Leggo i dati di carico del viaggio
    If Not DS.DA.QueryCommand(SQL_GET_LOAD_DATA, varBuffer, lngHowMany, IDViaggio) Then
        GoTo ErrHandler
    End If
    ' ---
    
    If InsertData Then
        ' --- Se non trovo i dati di carico determino quali prodotti devo
        ' --- caricare e inserisco i dati sul momento
        If lngHowMany = 0 Then
            If Not InsertLoadData(IDViaggio) Then
                GoTo ErrHandler
            End If
            If Not DS.DA.QueryCommand(SQL_GET_LOAD_DATA, varBuffer, lngHowMany, IDViaggio) Then
                GoTo ErrHandler
            End If
            If lngHowMany = 0 Then
                GoTo ErrHandler
            End If
        End If
        ' ---
    End If
    
    GetLoadData = varBuffer
    Exit Function
ErrHandler:
    GetLoadData = Empty
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


Public Function BeginSave(ByRef blnConcurErr As Boolean) As Boolean
    Dim blnSaveOk As Boolean
    Dim VVT As CViaggio
    
    Set VVT = New CViaggio
    With VVT.DS
        Set .DA = DS.DA
        Set .ErrH = DS.ErrH
    End With
    
    blnSaveOk = DS.BeginSave(SQL_UPDATE, SQL_GET_TS, blnConcurErr)
    
    VVT.DS.Fields("0").FindValue = DS.Fields("0").Value
    VVT.Find
    VVT.SetFlagTrasmissioneESys "N"
    
    BeginSave = True
    ' ---
End Function


Public Function InsertLoadData(IDViaggio As Variant) As Boolean
    On Error GoTo ErrHandler
    
'    Dim VVT As CViaggio
'
'    Set VVT = New CViaggio
'    With VVT.DS
'        Set .DA = DS.DA
'        Set .ErrH = DS.ErrH
'    End With
    
    If Not DS.DA.ExecuteCommand(SQL_INSERT_LOAD_DATA, IDViaggio) Then
        GoTo ErrHandler
    End If
    
    InsertLoadData = True
    
'    VVT.DS.Fields("0").FindValue = IDViaggio
'    VVT.Find
'    VVT.SetFlagTrasmissioneESys "N"
    
    Exit Function
    
ErrHandler:
    InsertLoadData = False
End Function

Public Function UpdateLoadData(QtaCarico As Variant, _
                               Den15 As Variant, _
                               DenAm As Variant, _
                               Rife1 As Variant, _
                               Rife2 As Variant, _
                               Rife3 As Variant, _
                               IDViaggio As Variant, _
                               IDProdotto As Variant, _
                               IDImballo As Variant, _
                               IDPosDog As Variant, _
                               QtaScarico As Variant) As Boolean
    On Error GoTo ErrHandler
    Dim VVT As CViaggio
    
    Set VVT = New CViaggio
    With VVT.DS
        Set .DA = DS.DA
        Set .ErrH = DS.ErrH
    End With
    
    If Not DS.DA.ExecuteCommand(SQL_UPDATE_LOAD_DATA, QtaCarico, Den15, DenAm, _
                                Rife1, Rife2, Rife3, QtaScarico, IDViaggio, IDProdotto, IDImballo, IDPosDog) Then
        GoTo ErrHandler
    End If
    
    VVT.DS.Fields("0").FindValue = IDViaggio
    VVT.Find
    VVT.SetFlagTrasmissioneESys "N"
   
    UpdateLoadData = True
    
    Exit Function
    
ErrHandler:
    UpdateLoadData = False
End Function

Private Sub Class_Initialize()
    SQL_FIND = "SELECT vcaidvvt, vcaidapr, vcaidaim, vcaidapd, " & _
               "       vcaqcakg, vcaqcalt, vcaqcapz, vcaqcacn, " & _
               "       vcarife1, vcarife2, vcarife3, vcaden15, " & _
               "       vcadenam, vcaqsclt, timestmp FROM viagvca0 WHERE "
    SQL_UPDATE = "UPDATE viagvca0 SET "
    SQL_INSERT = "INSERT INTO viagvca0 "
    SQL_GET_TS = "SELECT timestmp FROM viagvca0 WHERE "
    SQL_GET_LOAD_DATA = "SELECT vcaidapr, vcaidaim, vcaidapd, vcaqcalt, " & _
                        "       vcaden15, vcadenam, vcarife1, vcarife2, " & _
                        "       vcarife3, vcaqsclt " & _
                        "  FROM viagvca0 " & _
                        " WHERE vcaidvvt = ?"
    SQL_INSERT_LOAD_DATA = "INSERT INTO VIAGVCA0 (VCAIDVVT, VCAIDAPR, VCAIDAIM, VCAIDAPD) " & _
                           "SELECT DISTINCT VVPIDVVT, VVPIDAPR, VVPIDAIM, VVPIDAPD " & _
                           "  FROM VIAGVVP0 INNER JOIN VIAGVVC0 ON VVPIDVVT = VVCIDVVT AND VVPIDVVC = VVCIDVVC " & _
                           " WHERE VVPIDVVT=? AND VVPANNUL = 0 AND VVCANNUL = 0"
    SQL_UPDATE_LOAD_DATA = "UPDATE viagvca0 " & _
                           "   SET vcaqcalt = ?, vcaden15 = ?, vcadenam = ?, vcarife1 = ?, " & _
                       "           vcarife2 = ?, vcarife3 = ?, vcaqsclt = ? " & _
                       "     WHERE vcaidvvt = ? AND vcaidapr = ? AND vcaidaim = ? AND vcaidapd = ?"
                       
    SQL_DELETE_LOAD_DATA = "DELETE FROM VIAGVCA0 WHERE VCAIDVVT = ?"
    
    Set DS = New CDataStruct
    
    With DS.Fields
        .Add "Numero Viaggio", "viagvca0.vcaidvvt", 0, True, 7, "0", , True
        .Add "Prodotto", "viagvca0.vcaidapr", 0, True, 6, "1", , True
        With .Item("1")
            .ForeignTableEntity = "APR"
            .ForeignTable = "anagapr0"
            .ForeignTableFld = "apridapr"
            .ForeignTableDescField = "aprdesco"
            .ForeignTableDeleteFld = "aprannul"
        End With
        .Add "Imballo", "viagvca0.vcaidaim", 0, True, 3, "2", , True
        With .Item("2")
            .ForeignTableEntity = "AIM"
            .ForeignTable = "anagaim0"
            .ForeignTableFld = "aimidaim"
            .ForeignTableDescField = "aimdescr"
            .ForeignTableDeleteFld = "aimannul"
        End With
        .Add "Posizione Doganale", "viagvca0.vcaidapd", 0, True, 3, "3", , True
        With .Item("3")
            .ForeignTableEntity = "APD"
            .ForeignTable = "anagapd0"
            .ForeignTableFld = "apdidapd"
            .ForeignTableDescField = "apddescr"
            .ForeignTableDeleteFld = "apdannul"
        End With
        .Add "Quantita Trasportata Kg", "viagvca0.vcaqcakg", 1, False, 15, "4" ', , True
        .Add "Quantita Trasportata L", "viagvca0.vcaqcalt", 1, False, 15, "5" ', , True
        .Add "Quantita Trasportata Pz", "viagvca0.vcaqcapz", 1, False, 15, "6" ', , True
        .Add "Quantita Trasportata Cnf", "viagvca0.vcaqcacn", 1, False, 15, "7" ', , True
        .Add "Riferimento Esterno 1", "viagvca0.vcarife1", 0, False, 15, "8"
        .Add "Riferimento Esterno 2", "viagvca0.vcarife2", 0, False, 15, "9"
        .Add "Riferimento Esterno 3", "viagvca0.vcarife2", 0, False, 15, "10"
        .Add "Densità 15", "viagvca0.vcaden15", 1, False, 6, "11" ', , True
        .Add "Densità Ambiente", "viagvca0.vcadenam", 1, False, 6, "12" ', , True
        .Add "Quantita Scaricata L", "viagvca0.vcaqsclt", 1, False, 15, "13" ', , True
    End With
    With DS.KeyFields
        .Add "0", "1"
        .Add "1", "2"
        .Add "2", "3"
        .Add "3", "4"
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


Public Property Get SearchListFields() As Collection
    Set SearchListFields = mcolListFields
End Property

