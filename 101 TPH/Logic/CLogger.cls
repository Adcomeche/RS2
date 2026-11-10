VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Classe per la gestione del log

Private Const lsInfoId As Long = 0
Private Const lsDebugId As Long = 1
Private Const lsWarningId As Long = 2
Private Const lsErrorId As Long = 3

Private Const lsInfoDescription As String = "Info"
Private Const lsDebugDescription As String = "Debug"
Private Const lsWarningDescription As String = "Warning"
Private Const lsErrorDescription As String = "Error"

Private Const lsInfoStringId As String = "I"
Private Const lsDebugStringId As String = "D"
Private Const lsWarningStringId As String = "W"
Private Const lsErrorStringId As String = "E"

Private strLogFileName As String
Dim strLogStringsId(lsInfoId To lsErrorId) As String
Dim strLogDescriptions(lsInfoId To lsErrorId) As String
Dim blnLogEnabledLevels(lsInfoId To lsErrorId) As Boolean

Private lngSlitSize As Long

Public Property Let LogFileName(strLogFileNameNew As String)
    strLogFileName = strLogFileNameNew
End Property

Public Property Get LogFileName() As String
    LogFileName = strLogFileName
End Property

Public Sub enableLevel(strLogEnables As String)
    
    Dim lngLevelCounter As Long

    For lngLevelCounter = LBound(strLogStringsId) To UBound(strLogStringsId)
        If InStr(strLogEnables, strLogStringsId(lngLevelCounter)) > 0 Then
            blnLogEnabledLevels(lngLevelCounter) = True
        Else
            blnLogEnabledLevels(lngLevelCounter) = False
        End If
    Next lngLevelCounter
End Sub

Public Sub enableError(blnEnable As Boolean)
    blnLogEnabledLevels(lsErrorId) = blnEnable
End Sub

Public Sub enableWarning(blnEnable As Boolean)
    blnLogEnabledLevels(lsWarningId) = blnEnable
End Sub

Public Sub enableDebug(blnEnable As Boolean)
    blnLogEnabledLevels(lsDebugId) = blnEnable
End Sub

Public Sub enableInfo(blnEnable As Boolean)
    blnLogEnabledLevels(lsInfoId) = blnEnable
End Sub

Private Function GetFileSize(strFileName As String) As Long
On Error GoTo ErrHandler

    GetFileSize = FileSystem.FileLen(strLogFileName)

	Exit Function
ErrHandler:
    GetFileSize = 0
End Function

Public Sub log(lngSeverity As Long, strMessage As String)
On Error GoTo ErrHandler

    Dim strArchiveLogFileName As String
    
    If lngSeverity < lsInfoId Or lngSeverity > lsErrorId Then
        Exit Sub
    End If
    
    If Not blnLogEnabledLevels(lngSeverity) Then
        Exit Sub
    End If
    
    If IsEmpty(strLogFileName) Then
        Exit Sub
    End If
    
    If GetFileSize(strLogFileName) >= lngSlitSize * 1024 Then
        Dim lngDotIndex As Long
        
        If InStrRev(strLogFileName, ".") > 0 Then
            strArchiveLogFileName = Left(strLogFileName, InStrRev(strLogFileName, ".") - 1) & _
                Format(Now(), "yyyymmdd_HhNn") & Mid(strLogFileName, InStrRev(strLogFileName, "."))
        Else
            strArchiveLogFileName = strLogFileName & Format(Now(), "yyyymmdd_HhNn")
        End If

        FileSystem.FileCopy strLogFileName, strArchiveLogFileName
        FileSystem.Kill strLogFileName
    End If

    Open strLogFileName For Append As #1
    
    Print #1, Format(Now(), "dd/mm/yyyy Hh:Nn.Ss") & " [" & strLogStringsId(lngSeverity) _
                            ; "] - " & strMessage
    Close #1
    Exit Sub

ErrHandler:
    Close #1
End Sub

Public Sub logInfo(strMessage As String)
    log lsInfoId, strMessage
End Sub

Public Sub logWarning(strMessage As String)
    log lsWarningId, strMessage
End Sub

Public Sub logError(strMessage As String)
    log lsErrorId, strMessage
End Sub

Public Sub logDebug(strMessage As String)
    log lsDebugId, strMessage
End Sub

Private Sub Class_Initialize()

    strLogDescriptions(lsInfoId) = lsInfoDescription
    strLogDescriptions(lsDebugId) = lsDebugDescription
    strLogDescriptions(lsWarningId) = lsWarningDescription
    strLogDescriptions(lsErrorId) = lsErrorDescription
    
    strLogStringsId(lsInfoId) = lsInfoStringId
    strLogStringsId(lsDebugId) = lsDebugStringId
    strLogStringsId(lsWarningId) = lsWarnigStringId
    strLogStringsId(lsErrorId) = lsErrorStringId
    
    lngSlitSize = 1024

    enableError True
    enableWarning True
    enableInfo True
    enableDebug True

End Sub
