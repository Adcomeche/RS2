VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmSatellitareProcedure 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Procedure per l'aggiornamento dei dati satellitare"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11250
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7080
   ScaleWidth      =   11250
   Begin VB.Frame Frame1 
      Caption         =   " Parametri "
      Height          =   1095
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   11175
      Begin VB.Timer timerExecuteProcedure 
         Enabled         =   0   'False
         Left            =   4080
         Top             =   180
      End
      Begin MSMask.MaskEdBox mskTimerSeconds 
         Height          =   315
         Left            =   2880
         TabIndex        =   6
         Top             =   300
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   556
         _Version        =   393216
         ClipMode        =   1
         MaxLength       =   4
         Mask            =   "9999"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskFileNameTPH 
         Height          =   255
         Left            =   7920
         TabIndex        =   10
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFileNameEUT 
         Height          =   255
         Left            =   7920
         TabIndex        =   12
         Top             =   720
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskLastEvent 
         Height          =   315
         Left            =   2880
         TabIndex        =   14
         Top             =   720
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         ClipMode        =   1
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ultimo evento elaborato:"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   15
         Top             =   720
         Width           =   1995
      End
      Begin VB.Label Label1 
         Caption         =   "Percorso File EUT:"
         Height          =   255
         Index           =   0
         Left            =   6120
         TabIndex        =   13
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Percorso File TPH:"
         Height          =   255
         Index           =   14
         Left            =   6120
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblMinuti 
         Alignment       =   1  'Right Justify
         Caption         =   "Esecuzione delle procedure ogni                  secondi"
         Height          =   495
         Left            =   180
         TabIndex        =   5
         Top             =   360
         Width           =   3915
      End
   End
   Begin VB.TextBox textLog 
      CausesValidation=   0   'False
      Height          =   2955
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3240
      Width           =   11175
   End
   Begin VB.Frame Frame3 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   6240
      Width           =   11235
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   5
         Left            =   300
         Picture         =   "frmSatellitareProcedure.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   360
         Width           =   240
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Avvia"
         Height          =   375
         Left            =   180
         Picture         =   "frmSatellitareProcedure.frx":0102
         TabIndex        =   8
         Tag             =   "EXIT"
         Top             =   300
         Width           =   1575
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   9600
         Picture         =   "frmSatellitareProcedure.frx":0204
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   360
         Width           =   240
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   9480
         Picture         =   "frmSatellitareProcedure.frx":0306
         TabIndex        =   1
         Tag             =   "EXIT"
         Top             =   300
         Width           =   1575
      End
   End
   Begin TrueDBGrid60.TDBGrid gridProcedures 
      Height          =   1815
      Left            =   0
      OleObjectBlob   =   "frmSatellitareProcedure.frx":0408
      TabIndex        =   7
      Top             =   1320
      Width           =   11175
   End
End
Attribute VB_Name = "frmSatellitareProcedure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim XA As XArrayDB
Dim lngTimerSeconds As Long
Dim lngTimerTicks As Long
Dim strLogLevels As String

Private Const SECONDS_PER_TICK = 5
Private Const REGISTRY_KEY_NAME_EXECUTE_SECONDS = "ExecuteSecondsSatellitare"
Private Const REGISTRY_KEY_NAME_LOG_LEVELS = "LogSatellitareLevel"
Private Const REGISTRY_KEY_NAME_PATH_TPH = "PosFilePath_TPH"
Private Const REGISTRY_KEY_NAME_PATH_EUT = "PosFilePath_EUT"

Private mobjOutFileTPH As CTextFile
Private mobjOutFileEUT As CTextFile
Private mobjOutFile As CTextFile
Private Function Execute_Trasm_Posizioni_CEPSA() As Boolean
On Error GoTo ErrHandler
    
    Dim KVE As CSatellitareEvento
    Dim varMaxPos As Variant
    Dim varBufferPos As Variant
    Dim varBufferViaggio As Variant
    Dim varCounter As Variant
    Dim varRimorchio As Variant
    Dim strData As String
    Dim strOra As String
    Dim strMotrice As String
    Dim strIDDeposito As String
    Dim strProgrInvio As String
    Dim strFileName As String
    Dim strLine As String
    Dim i As Long
    Dim lngCountOK_TPH As Long
    Dim lngCountOK_EUT As Long
    Dim intFileFromNum As Integer
    Dim intFileToNum As Integer
    Dim VVT As CViaggio
    Dim AUT As CUnitaDiTrasporto
    Dim ADE As CDeposito
    Dim AAM As CAutomezzo
    
    Set VVT = New CViaggio
    With VVT.DS
        Set .DA = DataConnect.DA
        Set .ErrH = DataConnect.DA.ErrH
    End With
    Set AUT = New CUnitaDiTrasporto
    With AUT.DS
        Set .DA = DataConnect.DA
        Set .ErrH = DataConnect.DA.ErrH
    End With
    Set ADE = New CDeposito
    With ADE.DS
        Set .DA = DataConnect.DA
        Set .ErrH = DataConnect.DA.ErrH
    End With
    Set AAM = New CAutomezzo
    With AAM.DS
        Set .DA = DataConnect.DA
        Set .ErrH = DataConnect.DA.ErrH
    End With
    
     '--SM -JGFichero para la Detección de ejecución de proccesos
    Dim ret As Boolean
    Dim x As CTextFile
    Set x = New CTextFile
    ret = x.OpenFile("R:\Resoil2\Alarmas\GIS.txt", 1)
    ret = x.CloseFile()
    
    
    Set KVE = New CSatellitareEvento

    With KVE
        Set .DS.DA = DataConnect.DA
        Set .DS.ErrH = DataConnect.DA.ErrH
        
        .Logger.LogFileName = gmstrServerLogFileName
        .Logger.enableLevel (strLogLevels)
    
        LogMessage "Avvio procedura Execute_Trasm_Posizioni_CEPSA." & vbCrLf

        XA.value(1, 1) = "In esecuzione."
        gridProcedures.RefetchCol 1
        
        varMaxPos = .SAT_Get_Max_Pos
        If IsEmpty(varMaxPos) Then
            GoTo ErrHandler
        End If
        
        varBufferPos = .SAT_Get_Pos_List(varMaxPos)
        If Not IsEmpty(varBufferPos) Then
            If Not mobjOutFileTPH.OpenFile("tphpos.txt", 1) Then
                GoTo ErrHandler
            End If
            If Not mobjOutFileEUT.OpenFile("eutpos.txt", 1) Then
                GoTo ErrHandler
            End If
            strData = Format(Date, "yyyymmdd")
            strOra = Replace(Format(Time, "hh:mm:ss"), ".", ":")
            For i = 0 To UBound(varBufferPos, 2)
                strMotrice = varBufferPos(0, i)
                
                If VVT.GetUltimoViaggioCaricato(strMotrice, varBufferPos(1, i), varBufferPos(2, i), varBufferViaggio) Then
                    Select Case varBufferViaggio(1, 0)
                        Case "004"
                            Set mobjOutFile = mobjOutFileEUT
                        Case "624"
                            Set mobjOutFile = mobjOutFileTPH
                    End Select
                    With ADE.DS
                        .Fields.ResetFields
                        .Fields("0").FindValue = varBufferViaggio(4, 0)
                        If (Not ADE.Find) Or (.HowMany = 0) Then
                            GoTo ErrHandler
                        End If
                        strIDDeposito = .MultiSoc.DS.GetFirmCode(ID_SOC_CLH)
                    End With
                    varRimorchio = varBufferViaggio(5, 0)
                    
                    If Left(strMotrice, 2) <> "ES" Then
                        AAM.DS.Fields.ResetFields
                        AAM.DS.Field("Targa").FindValue = strMotrice
                        AAM.Find
                        If Len(AAM.DS.Field("Flag Nazionalità").value) > 0 Then
                            strMotrice = UCase(AAM.DS.Field("Flag Nazionalità").value) & strMotrice
                        Else
                            strMotrice = "ES" & strMotrice
                        End If
                    End If
                    strMotrice = Replace(Replace(strMotrice, " ", ""), "-", "")
                    
                    If Len(varRimorchio) > 0 Then
                        AAM.DS.Fields.ResetFields
                        AAM.DS.Field("Targa").FindValue = varRimorchio
                        AAM.Find
                        If Left(varRimorchio, 2) <> "ES" Then
                            If Len(AAM.DS.Field("Flag Nazionalità").value) > 0 Then
                                varRimorchio = UCase(AAM.DS.Field("Flag Nazionalità").value) & varRimorchio
                            Else
                                varRimorchio = "ES" & varRimorchio
                            End If
                        End If
                        varRimorchio = Replace(Replace(varRimorchio, " ", ""), "-", "")
                    End If
                    
                    With mobjOutFile
                        .ClearFields
                        .Item("1").value = Format(varBufferViaggio(1, 0), "0000")
                        .Item("3").value = strData
                        .Item("4").value = strOra
                        .Item("6").value = varBufferViaggio(3, 0)
                        .Item("7").value = strIDDeposito
                        .Item("8").value = strMotrice
                        .Item("9").value = varRimorchio
                        .Item("12").value = Replace(Format(varBufferPos(3, i), "00.00000"), ".", ",")
                        .Item("13").value = Replace(Format(varBufferPos(4, i), "00.000000"), ".", ",")
                        .Item("14").value = Format(varBufferPos(1, i), "yyyymmdd")
                        .Item("15").value = Replace(Format(varBufferPos(2, i), "hh:mm:ss"), ".", ":")
                        If Not .WriteLine Then
                            GoTo ErrHandler
                        End If
                        Select Case varBufferViaggio(1, 0)
                            Case "004"
                                lngCountOK_EUT = lngCountOK_EUT + 1
                            Case "624"
                                lngCountOK_TPH = lngCountOK_TPH + 1
                        End Select
                    End With
                End If
            Next i
            If Not mobjOutFileTPH.CloseFile Then
                GoTo ErrHandler
            End If
            If Not mobjOutFileEUT.CloseFile Then
                GoTo ErrHandler
            End If
        End If
        
        If Not .SAT_Update_Pos_Status(varMaxPos) Then
            GoTo ErrHandler
        End If
        
    End With
    
    If lngCountOK_TPH > 0 Then
        With KVE.DS.DA
            varCounter = .GetCounter("anaganu0", "anuprogr", "anuidanu", "030", 0)
            If varCounter = -1 Then
                GoTo ErrHandler
            End If
            If varCounter > 9999999999# Then
                varCounter = 1
            End If
        End With
        
        strProgrInvio = Format(varCounter, String(10, "0"))
        strFileName = mskFileNameTPH & "\GIS_" & strProgrInvio & ".DAT"
        intFileFromNum = FreeFile
        
        Open "tphpos.txt" For Input Access Read As #intFileFromNum
        intFileToNum = FreeFile
        Open strFileName For Output Access Write As #intFileToNum
        Do While Not EOF(intFileFromNum)
            Line Input #intFileFromNum, strLine
            Mid(strLine, 31, 5) = Format(lngCountOK_TPH, "00000")
            Mid(strLine, 5, 10) = strProgrInvio
            Print #intFileToNum, strLine
        Loop
        Close #intFileFromNum
        Close #intFileToNum
        
'        FileCopy "tphpos.txt", strFileName
    End If
    If lngCountOK_EUT > 0 Then
        With KVE.DS.DA
            varCounter = .GetCounter("anaganu0", "anuprogr", "anuidanu", "031", 0)
            If varCounter = -1 Then
                GoTo ErrHandler
            End If
            If varCounter > 9999999999# Then
                varCounter = 1
            End If
        End With
        
        strProgrInvio = Format(varCounter, String(10, "0"))
        strFileName = mskFileNameEUT & "\GIS_" & strProgrInvio & ".DAT"
        intFileFromNum = FreeFile
        
        Open "eutpos.txt" For Input Access Read As #intFileFromNum
        intFileToNum = FreeFile
        Open strFileName For Output Access Write As #intFileToNum
        Do While Not EOF(intFileFromNum)
            Line Input #intFileFromNum, strLine
            Mid(strLine, 31, 5) = Format(lngCountOK_EUT, "00000")
            Mid(strLine, 5, 10) = strProgrInvio
            Print #intFileToNum, strLine
        Loop
        Close #intFileFromNum
        Close #intFileToNum
'        FileCopy "eutpos.txt", strFileName
    End If
    
    XA.value(1, 1) = "In attesa."
    gridProcedures.RefetchCol 1

    Execute_Trasm_Posizioni_CEPSA = True
    Set KVE = Nothing
    Exit Function

ErrHandler:
    Set KVE = Nothing
    mobjOutFile.CloseFile
    mobjOutFileEUT.CloseFile
    mobjOutFileTPH.CloseFile
    XA.value(1, 1) = "In errore."
    gridProcedures.RefetchCol 1
End Function

Private Sub Command1_Click()
    If lngTimerTicks = -1 Then
        StartTimer
    Else
        StopTimer
    End If
End Sub
Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim lngItemsCounter As Long
    Set XA = New XArrayDB

    XA.ReDim 0, 2, 0, 2
    
    XA.value(0, 0) = "Aggiornamento posizioni Oesia."
    XA.value(0, 1) = "In attesa"
    XA.value(0, 2) = "Esegui"
    
'    XA.value(1, 0) = "Aggiornamento posizioni Deusto."
'    XA.value(1, 1) = "In attesa"
'    XA.value(1, 2) = "Esegui"
    
    XA.value(1, 0) = "Invio posizioni a CEPSA."
    XA.value(1, 1) = "In attesa"
    XA.value(1, 2) = "Esegui"
    
'    XA.value(2, 0) = "Aggiornamento messaggi ricevuti KNOSOS."
'    XA.value(2, 1) = "In attesa"
'    XA.value(2, 2) = "Esegui"

    gridProcedures.Array = XA
    gridProcedures.ReBind

    SetTimerInterval CLng(gmRegistry.GetSetting("Options", REGISTRY_KEY_NAME_EXECUTE_SECONDS, "15", True)), True
    StartTimer

    strLogLevels = gmRegistry.GetSetting("Options", REGISTRY_KEY_NAME_LOG_LEVELS, "IDEW", True)
    Me.mskFileNameEUT = gmRegistry.GetSetting("Options", REGISTRY_KEY_NAME_PATH_EUT, "", True)
    Me.mskFileNameTPH = gmRegistry.GetSetting("Options", REGISTRY_KEY_NAME_PATH_TPH, "", True)
    
    Set mobjOutFile = New CTextFile
    With mobjOutFile
        .Add "Codice Trasportatore CLH", 4, "1"
        .Add "Progressivo invio", 10, "2", , fftNumerico
        .Add "Data generazione", 8, "3"
        .Add "Ora generazione", 8, "4"
        .Add "Numero di records", 5, "5", , fftNumerico
        .Add "Badge", 11, "6"
        .Add "Codice CLH Deposito", 4, "7"
        .Add "Targa 1", 18, "8"
        .Add "Targa 2", 18, "9"
        .Add "Targa 3", 18, "10"
        .Add "Tipo coordinata", 2, "11", "01"
        .Add "Longitudine (X)", 9, "12"
        .Add "Latitudine (Y)", 9, "13"
        .Add "Data posizione", 8, "14"
        .Add "Ora posizione", 8, "15"
    End With
    Set mobjOutFileTPH = New CTextFile
    With mobjOutFileTPH
        .Add "Codice Trasportatore CLH", 4, "1"
        .Add "Progressivo invio", 10, "2", , fftNumerico
        .Add "Data generazione", 8, "3"
        .Add "Ora generazione", 8, "4"
        .Add "Numero di records", 5, "5", , fftNumerico
        .Add "Badge", 11, "6"
        .Add "Codice CLH Deposito", 4, "7"
        .Add "Targa 1", 18, "8"
        .Add "Targa 2", 18, "9"
        .Add "Targa 3", 18, "10"
        .Add "Tipo coordinata", 2, "11", "01"
        .Add "Longitudine (X)", 9, "12"
        .Add "Latitudine (Y)", 9, "13"
        .Add "Data posizione", 8, "14"
        .Add "Ora posizione", 8, "15"
    End With
    Set mobjOutFileEUT = New CTextFile
    With mobjOutFileEUT
        .Add "Codice Trasportatore CLH", 4, "1"
        .Add "Progressivo invio", 10, "2", , fftNumerico
        .Add "Data generazione", 8, "3"
        .Add "Ora generazione", 8, "4"
        .Add "Numero di records", 5, "5", , fftNumerico
        .Add "Badge", 11, "6"
        .Add "Codice CLH Deposito", 4, "7"
        .Add "Targa 1", 18, "8"
        .Add "Targa 2", 18, "9"
        .Add "Targa 3", 18, "10"
        .Add "Tipo coordinata", 2, "11", "01"
        .Add "Longitudine (X)", 9, "12"
        .Add "Latitudine (Y)", 9, "13"
        .Add "Data posizione", 8, "14"
        .Add "Ora posizione", 8, "15"
    End With
    
    mskLastEvent.Text = gmRegistry.GetSetting("Options", "LastEvent_Oesia", "0", True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    gmRegistry.SaveSetting "Options", REGISTRY_KEY_NAME_PATH_TPH, Me.mskFileNameTPH, True
    gmRegistry.SaveSetting "Options", REGISTRY_KEY_NAME_PATH_EUT, Me.mskFileNameEUT, True
    gmRegistry.SaveSetting "Options", "LastEvent_Oesia", mskLastEvent.Text, True
    gmRegistry.SaveSetting "Options", REGISTRY_KEY_NAME_EXECUTE_SECONDS, Me.mskTimerSeconds, True
    
    Set XA = Nothing

    Set frmSatellitareProcedure = Nothing
End Sub

Private Sub gridProcedures_ButtonClick(ByVal ColIndex As Integer)
    If ColIndex = 2 Then
        Select Case gridProcedures.bookmark
            Case 0
                Execute_ScanEvents_Oesia
            
            Case 1
                Execute_Trasm_Posizioni_CEPSA
                
        End Select
    End If
End Sub
Private Sub LogMessage(message As String)
    textLog.Text = textLog.Text & Format(Now(), "Hh:Nn.Ss") & " - " & message
    textLog.SelStart = Len(textLog.Text)
End Sub

Private Function ExecuteProcedure()
    
    If Len(textLog.Text) > 10000 Then
        textLog.Text = ""
    End If
    
    ' Esegue la procedura di Scan Events Oesia
    Execute_ScanEvents_Oesia
  
    
    
    ' Esegue la procedura di Scan Events Deusto
  '  Execute_ScanEvents_Deusto
    
    
    ' Trasmette le posizioni a Cepsa
 '   Execute_Trasm_Posizioni_CEPSA
   
    
'    ' Esegue la procedura di Scan Messaggi Trasmessi
'    Execute_ScanMessageSent_KNOSOS
    
'    ' Esegue la procedura di Scan Messaggi Ricevuti
'    Execute_ScanMessageReceived_KNOSOS

End Function

Private Function Execute_ScanEvents_Oesia() As Boolean
    Dim varLastID As Variant
    Dim KVE As CSatellitareEvento

    On Error GoTo ErrHandler
    
    varLastID = mskLastEvent.Text
    
      '--SM -JGFichero para la Detección de ejecución de proccesos
    Dim ret As Boolean
    Dim x As CTextFile
    Set x = New CTextFile
    ret = x.OpenFile("R:\Resoil2\Alarmas\DISTVEHI.txt", 1)
    ret = x.CloseFile()
    
    
    Set KVE = New CSatellitareEvento

    With KVE
        Set .DS.DA = DataConnect.DA
        Set .DS.ErrH = DataConnect.DA.ErrH
        
        KVE.Logger.LogFileName = gmstrServerLogFileName
        KVE.Logger.enableLevel (strLogLevels)
    
        LogMessage "Avvio procedura Scan_Events_OESIA." & vbCrLf

        XA.value(0, 1) = "In esecuzione."
        gridProcedures.RefetchCol 1
        
        If Not .ScanEvents_OESIA(varLastID) Then
            GoTo ErrHandler
        End If
    
    End With

    XA.value(0, 1) = "In attesa."
    gridProcedures.RefetchCol 1

    Execute_ScanEvents_Oesia = True
    Set KVE = Nothing
    GoTo EndProc

ErrHandler:
    Set KVE = Nothing
    
    XA.value(0, 1) = "In errore."
    gridProcedures.RefetchCol 1
    ret = x.OpenFile("R:\Resoil2\Alarmas\RESOILSQL.txt", 1)
    ret = x.CloseFile()

    
EndProc:
    mskLastEvent.Text = varLastID
    gmRegistry.SaveSetting "Options", "LastEvent_Oesia", varLastID, True
End Function

Private Function Execute_ScanEvents_Deusto() As Boolean
On Error GoTo ErrHandler
    
    Dim KVE As CSatellitareEvento

    Set KVE = New CSatellitareEvento

'--SM -JGFichero para la Detección de ejecución de proccesos
    Dim ret As Boolean
    Dim x As CTextFile
    Set x = New CTextFile
    ret = x.OpenFile("R:\Resoil2\Alarmas\POSVEHI.txt", 1)
    ret = x.CloseFile()
    
    With KVE
        Set .DS.DA = DataConnect.DA
        Set .DS.ErrH = DataConnect.DA.ErrH
        
        KVE.Logger.LogFileName = gmstrServerLogFileName
        KVE.Logger.enableLevel (strLogLevels)
    
        LogMessage "Avvio procedura Scan_Events_DEUSTO." & vbCrLf

        XA.value(1, 1) = "In esecuzione."
        gridProcedures.RefetchCol 1
        
        If Not .ScanEvents_DEUSTO() Then
            GoTo ErrHandler
        End If
    
    End With

    XA.value(1, 1) = "In attesa."
    gridProcedures.RefetchCol 1

    Execute_ScanEvents_Deusto = True
    Set KVE = Nothing
    Exit Function

ErrHandler:
    Set KVE = Nothing
    
    XA.value(1, 1) = "In errore."
    gridProcedures.RefetchCol 1
End Function

Private Function Execute_ScanMessageSent_KNOSOS()
On Error GoTo ErrHandler

    Dim SME As CSatellitareMessaggio

    Set SME = New CSatellitareMessaggio

    SME.Logger.LogFileName = gmstrServerLogFileName
    
    XA.value(1, 1) = "In esecuzione."
    gridProcedures.RefetchCol 1
    
    With SME
        Set .DS.DA = DataConnect.DA
        Set .DS.ErrH = DataConnect.DA.ErrH

        LogMessage "Avvio procedura Scan_MessageSent_KNOSOS." & vbCrLf
        
        If Not .Scan_MessageSent_KNOSOS() Then
            GoTo ErrHandler
        End If

    End With
    
    XA.value(1, 1) = "In attesa."
    gridProcedures.RefetchCol 1

    Set SME = Nothing
    Execute_ScanMessageSent_KNOSOS = True
    Exit Function

ErrHandler:
    Set SME = Nothing
    XA.value(1, 1) = "In errore."
    gridProcedures.RefetchCol 1
    
    Execute_ScanMessageSent_KNOSOS = False
End Function
Private Function Execute_ScanMessageReceived_KNOSOS()
On Error GoTo ErrHandler

    Dim SME As CSatellitareMessaggio

    Set SME = New CSatellitareMessaggio
    SME.Logger.LogFileName = gmstrServerLogFileName
    
    XA.value(2, 1) = "In esecuzione."
    gridProcedures.RefetchCol 1
    
    With SME
        Set .DS.DA = DataConnect.DA
        Set .DS.ErrH = DataConnect.DA.ErrH

        LogMessage "Avvio procedura ScanMessageReceived_KNOSOS." & vbCrLf
        
        If Not .Scan_MessageReceived_KNOSOS() Then
            GoTo ErrHandler
        End If

    End With
    
    XA.value(2, 1) = "In attesa."
    gridProcedures.RefetchCol 1

    Set SME = Nothing
    Execute_ScanMessageReceived_KNOSOS = True
    Exit Function

ErrHandler:
    Set SME = Nothing
    XA.value(2, 1) = "In errore."
    gridProcedures.RefetchCol 1
    
    Execute_ScanMessageReceived_KNOSOS = False
End Function

Private Sub mskTimerSeconds_Validate(Cancel As Boolean)
On Error GoTo ErrHandler
    
    Dim blnTimerEnabled As Boolean
    Dim blnDummy As Boolean

    If CLng(mskTimerSeconds.FormattedText) > 0 Then
        blnTimerEnabled = IIf(lngTimerTicks = -1, False, True)
        
        StopTimer
        SetTimerInterval CLng(mskTimerSeconds.FormattedText), False
        blnDummy = gmRegistry.SaveSetting("Options", REGISTRY_KEY_NAME_EXECUTE_SECONDS, CStr(CLng(mskTimerSeconds.FormattedText)), True)
        
        If blnTimerEnabled Then
            StartTimer
        End If
        Exit Sub
    Else
        If CLng(mskTimerSeconds.FormattedText) = 0 Then
            StopTimer
        End If
    End If

ErrHandler:
        Cancel = True

End Sub
Private Sub timerExecuteProcedure_Timer()
On Error GoTo ErrHandler

    timerExecuteProcedure.Enabled = False
    lngTimerTicks = lngTimerTicks + 1

    If lngTimerTicks * SECONDS_PER_TICK >= lngTimerSeconds Then
        ExecuteProcedure
        StartTimer
        Exit Sub
    End If

    timerExecuteProcedure.Enabled = True
    Exit Sub

ErrHandler:
    timerExecuteProcedure.Enabled = True
End Sub
Private Sub SetTimerInterval(lngSeconds As Long, blnRefreshControl As Boolean)
    StopTimer
    lngTimerSeconds = lngSeconds
    If lngTimerSeconds > 0 And lngTimerSeconds < 10 Then
        lngTimerSeconds = 10
    End If
    If blnRefreshControl = True Then
        mskTimerSeconds.Text = CStr(lngTimerSeconds) + Space(mskTimerSeconds.MaxLength - Len(CStr(lngTimerSeconds)))
    End If
    StartTimer
End Sub
Private Sub StartTimer()
    If lngTimerSeconds = 0 Then
        Exit Sub
    End If

    lngTimerTicks = 0
    timerExecuteProcedure.Interval = SECONDS_PER_TICK * 1000
    Command1.Caption = "Arresta"
    timerExecuteProcedure.Enabled = True
End Sub
Private Sub StopTimer()
    timerExecuteProcedure.Enabled = False
    lngTimerTicks = -1
    Command1.Caption = "Avvia"
End Sub


