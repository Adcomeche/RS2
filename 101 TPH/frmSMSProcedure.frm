VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmSMSProcedure 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Procedure per l'aggiornamento dei dati da SMS"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11430
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7125
   ScaleWidth      =   11430
   Begin VB.Frame Frame1 
      Caption         =   " Parametri "
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   11175
      Begin VB.CheckBox chkpalm 
         Caption         =   "Recoger Eventos palm"
         Height          =   255
         Left            =   8040
         TabIndex        =   12
         Top             =   360
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.Timer timerExecuteProcedure 
         Enabled         =   0   'False
         Left            =   10620
         Top             =   180
      End
      Begin MSMask.MaskEdBox mskTimerSeconds 
         Height          =   315
         Left            =   2640
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
      Begin MSMask.MaskEdBox mskLastSMS 
         Height          =   315
         Left            =   5880
         TabIndex        =   10
         Top             =   360
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         ClipMode        =   1
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         Caption         =   "Ultimo SMS elaborato:"
         Height          =   255
         Left            =   4080
         TabIndex        =   11
         Top             =   360
         Width           =   1755
      End
      Begin VB.Label lblMinuti 
         Caption         =   "Esecuzione delle procedure ogni                  secondi"
         Height          =   255
         Left            =   180
         TabIndex        =   5
         Top             =   360
         Width           =   3915
      End
   End
   Begin VB.TextBox textLog 
      CausesValidation=   0   'False
      Height          =   3075
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3000
      Width           =   11175
   End
   Begin VB.Frame Frame3 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   6180
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
         Picture         =   "frmSMSProcedure.frx":0000
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
         Picture         =   "frmSMSProcedure.frx":0102
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
         Picture         =   "frmSMSProcedure.frx":0204
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
         Picture         =   "frmSMSProcedure.frx":0306
         TabIndex        =   1
         Tag             =   "EXIT"
         Top             =   300
         Width           =   1575
      End
   End
   Begin TrueDBGrid60.TDBGrid gridProcedures 
      Height          =   1815
      Left            =   120
      OleObjectBlob   =   "frmSMSProcedure.frx":0408
      TabIndex        =   7
      Top             =   1080
      Width           =   11175
   End
End
Attribute VB_Name = "frmSMSProcedure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim XA As XArrayDB
Dim lngTimerSeconds As Long
Dim lngTimerTicks As Long
Dim strLogLevels As String

Private mblnRunOk As Boolean

Private Const SECONDS_PER_TICK = 5
Private Const REGISTRY_KEY_NAME_EXECUTE_SECONDS = "ExecuteSecondsSMS"
Private Const REGISTRY_KEY_NAME_LOG_LEVELS = "LogSMSLevel"
Private Const REGISTRY_KEY_NAME_LAST_SMS = "LastSMSID"

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
    
    XA.value(0, 0) = "Aggiornamento eventi SMS."
    XA.value(0, 1) = "In attesa"
    XA.value(0, 2) = "Esegui"
    
'    XA.value(1, 0) = "Procedimiento de captura de eventos Palm"
'    XA.value(1, 1) = "A la Espera"
'    XA.value(1, 2) = "Ejecutar"

'    XA.value(2, 0) = "Aggiornamento messaggi ricevuti KNOSOS."
'    XA.value(2, 1) = "In attesa"
'    XA.value(2, 2) = "Esegui"

    gridProcedures.Array = XA
    gridProcedures.ReBind
    
    mblnRunOk = True
    
    SetTimerInterval CLng(gmRegistry.GetSetting("Options", REGISTRY_KEY_NAME_EXECUTE_SECONDS, "15", True)), True
    StartTimer

    strLogLevels = gmRegistry.GetSetting("Options", REGISTRY_KEY_NAME_LOG_LEVELS, "IDEW", True)
    mskLastSMS = gmRegistry.GetSetting("Options", REGISTRY_KEY_NAME_LAST_SMS, "0", True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set XA = Nothing
   gmRegistry.SaveSetting "Options", REGISTRY_KEY_NAME_LAST_SMS, mskLastSMS.Text, True
   
   Set frmSMSProcedure = Nothing
End Sub
Private Sub gridProcedures_ButtonClick(ByVal ColIndex As Integer)
    If ColIndex = 2 Then
        Select Case gridProcedures.bookmark
            Case 0
                Execute_ScanEvents
            
            Case 1
               ' Execute_ScanEventsPalm
            
            Case 2
               ' Execute_ScanMessageReceived_KNOSOS
        End Select
    End If

End Sub
Private Sub LogMessage(message As String)
    textLog.Text = textLog.Text & Format(Now(), "Hh:Nn.Ss") & " - " & message
    textLog.SelStart = Len(textLog.Text)
End Sub

Private Function ExecuteProcedure()

    Dim palmpres As Boolean
    
    
    If Len(textLog.Text) > 10000 Then
        textLog.Text = ""
    End If
    
    
 
     ' Esegue la procedura di Scan Events
    Execute_ScanEvents
       
    
'    ' Esegue la procedura di Scan Messaggi Trasmessi
    
'    palmpres = (chkpalm = vbChecked)
'
'    If palmpres Then
'
'        Execute_ScanEventsPalm
'    End If
    
'    ' Esegue la procedura di Scan Messaggi Ricevuti
'    Execute_ScanMessageReceived_KNOSOS

End Function

Private Function Execute_ScanEvents() As Boolean
On Error GoTo ErrHandler
    Dim varLastID As Variant
    Dim varMinuti As Variant
    
    Dim KVE As CSatellitareEvento
    
    varLastID = mskLastSMS.Text
    
    '--SM -JG Detección de ejecución de proccesos
    Dim ret As Boolean
    Dim x As CTextFile
    Set x = New CTextFile
    ret = x.OpenFile("R:\Resoil2\Alarmas\SMS.txt", 1)
    ret = x.CloseFile()
    
    
    Set KVE = New CSatellitareEvento
    With KVE
        Set .DS.DA = DataConnect.DA
        Set .DS.ErrH = DataConnect.DA.ErrH
        
        KVE.Logger.LogFileName = gmstrServerLogFileName
        KVE.Logger.enableLevel (strLogLevels)
    
        XA.value(0, 1) = "In esecuzione."
        gridProcedures.RefetchCol 1
        
'        varMinuti = Format(Time, "nn")
'        Select Case varMinuti
'            Case "00", "15", "30", "45"

'        Select Case varMinuti Mod 10
'            Case 0
                If mblnRunOk Then
                    '-- DP 19/05/08
                    '-- "AggiornaOrari_SMS" non serve più, si fa tutto da "ScanEvents_SMS"
'                    LogMessage "Avvio procedura AggiornaOrari_SMS." & vbCrLf
'                    If Not KVE.AggiornaOrari_SMS(Date, Time) Then
'                        GoTo ErrHandler
'                    End If
                    
                    mblnRunOk = False
                    LogMessage "Avvio procedura ScanEvents_SMS." & vbCrLf

                    If Not KVE.ScanEvents_SMS(varLastID) Then
                        GoTo ErrHandler
                    End If
                    
                    mblnRunOk = True
                End If
                
'            Case Else
'                mblnRunOk = True
'
'        End Select
        
    End With

    XA.value(0, 1) = "In attesa."
    gridProcedures.RefetchCol 1

    Execute_ScanEvents = True
    Set KVE = Nothing
    GoTo EndProc

ErrHandler:
    Set KVE = Nothing
    
    mblnRunOk = True
    XA.value(0, 1) = "In errore."
    gridProcedures.RefetchCol 1
    
EndProc:
    mskLastSMS.Text = varLastID
    gmRegistry.SaveSetting "Options", REGISTRY_KEY_NAME_LAST_SMS, varLastID, True
End Function
Private Function Execute_ScanEventsPalm() As Boolean
On Error GoTo ErrHandler
    Dim varLastID As Variant
    Dim varMinuti As Variant

    Dim KVE As CSatellitareEvento
    
    
 '--SM -JG Detección de ejecución de proccesos
        Dim ret As Boolean
        Dim x As CTextFile
        Set x = New CTextFile
        ret = x.OpenFile("R:\Resoil2\Alarmas\PALM.txt", 1)
        ret = x.CloseFile()
        

    Set KVE = New CSatellitareEvento
    With KVE
        Set .DS.DA = DataConnect.DA
        Set .DS.ErrH = DataConnect.DA.ErrH

'        KVE.Logger.LogFileName = gmstrServerLogFileName
'        KVE.Logger.enableLevel (strLogLevels)

        XA.value(1, 1) = "A la espera"
        gridProcedures.RefetchCol 1

                If mblnRunOk Then
                    mblnRunOk = False
                    LogMessage "Inicio poceso de escaneo eventos de Palms: " & vbCrLf

                    If Not KVE.ScanEvents_Palm Then
                        GoTo ErrHandler
                    End If

                    mblnRunOk = True
                End If

    End With

    XA.value(1, 1) = "En espera"
    gridProcedures.RefetchCol 1

    Execute_ScanEventsPalm = True
    Set KVE = Nothing
    GoTo EndProc

ErrHandler:
    Set KVE = Nothing

    mblnRunOk = True
    XA.value(1, 1) = "Error"
    gridProcedures.RefetchCol 1

EndProc:


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


