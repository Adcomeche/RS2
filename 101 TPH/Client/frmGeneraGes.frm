VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmGeneraGes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generacion de ficheros Ges"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmGeneraGes.frx":0000
   ScaleHeight     =   8760
   ScaleWidth      =   11205
   Begin VB.Frame Frame3 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   8
      Top             =   7920
      Width           =   11235
      Begin VB.CommandButton btn_pararev 
         Caption         =   "Activar/Desactivar Parar Eventos"
         Height          =   375
         Left            =   3000
         TabIndex        =   14
         Top             =   240
         Width           =   4095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   9480
         Picture         =   "frmGeneraGes.frx":0342
         TabIndex        =   12
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
         Picture         =   "frmGeneraGes.frx":0444
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   360
         Width           =   240
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Avvia"
         Height          =   375
         Left            =   180
         Picture         =   "frmGeneraGes.frx":0546
         TabIndex        =   10
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
         Index           =   5
         Left            =   300
         Picture         =   "frmGeneraGes.frx":0648
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   360
         Width           =   240
      End
   End
   Begin VB.TextBox textLog 
      CausesValidation=   0   'False
      Height          =   3075
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   4800
      Width           =   11175
   End
   Begin VB.Frame Frame1 
      Caption         =   " Parametri "
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      Begin VB.CheckBox Chkdesp 
         Caption         =   "Captura desprecintados"
         Height          =   255
         Left            =   8760
         TabIndex        =   19
         Top             =   120
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.Timer timerExecuteProcedure 
         Enabled         =   0   'False
         Left            =   4200
         Top             =   480
      End
      Begin MSMask.MaskEdBox mskTimerSeconds 
         Height          =   315
         Left            =   2880
         TabIndex        =   1
         Top             =   120
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
         TabIndex        =   2
         Top             =   360
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFileNameEUT 
         Height          =   255
         Left            =   7920
         TabIndex        =   3
         Top             =   720
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskLastSMS 
         Height          =   315
         Left            =   3120
         TabIndex        =   15
         Top             =   480
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         _Version        =   393216
         ClipMode        =   1
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskEdUnidad 
         Height          =   255
         Left            =   5640
         TabIndex        =   17
         Top             =   720
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdHoras 
         Height          =   315
         Left            =   3120
         TabIndex        =   20
         Top             =   720
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         _Version        =   393216
         ClipMode        =   1
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         Caption         =   "Margen Horas reenvio (Blanco no reenvia)"
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   21
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "Unidad Red"
         Height          =   255
         Index           =   2
         Left            =   4560
         TabIndex        =   18
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Ultimo SMS procesado"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   16
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblMinuti 
         Alignment       =   1  'Right Justify
         Caption         =   "Esecuzione delle procedure ogni                  secondi"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   3915
      End
      Begin VB.Label Label1 
         Caption         =   "Fichero GES TPH"
         Height          =   255
         Index           =   14
         Left            =   6120
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Fichero GES Eutrasur"
         Height          =   255
         Index           =   0
         Left            =   6120
         TabIndex        =   4
         Top             =   720
         Width           =   1695
      End
   End
   Begin TrueDBGrid60.TDBGrid gridProcedures 
      Height          =   3615
      Left            =   0
      OleObjectBlob   =   "frmGeneraGes.frx":074A
      TabIndex        =   13
      Top             =   1200
      Width           =   11175
   End
End
Attribute VB_Name = "frmGeneraGes"
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
Private Const REGISTRY_KEY_NAME_EXECUTE_SECONDS = "ExecuteSecondsSatellitare"
Private Const REGISTRY_KEY_NAME_LOG_LEVELS = "LogSatellitareLevel"
Private Const REGISTRY_KEY_NAME_PATH_TPH = "PosFilePath_TPH"
Private Const REGISTRY_KEY_NAME_PATH_EUT = "PosFilePath_EUT"
Private Const REGISTRY_KEY_NAME_LAST_SMS = "LastSMSID"
Private Const REGISTRY_KEY_NAME_UNIDAD = "Unidadresoil"
Private Const REGISTRY_KEY_NAME_HORAS = "Horasreenvio"

Private mobjOutFileTPH As CTextFile
Private mobjOutFileEUT As CTextFile
Private mobjOutFile As CTextFile
Private mobjFileAlb As CTextFile

'SM
Public WithEvents DS As CDataStruct
Attribute DS.VB_VarHelpID = -1
Public WithEvents DataClass As CViaggioSelezione
Attribute DataClass.VB_VarHelpID = -1
Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1
Private ProcesoEnCurso As Boolean



Private Function Actualiza_Horas()
Dim DA As CDataAccess
Set DA = DataConnect.DA
Dim SQL_GEN_EV As String
Dim strIDViaggio, strIDConsegna As String



SQL_GEN_EV = "INSERT INTO EventosMBFlotasManual (CodigoServicio,CodigoTarea,CodigoEvento,FechaEvento,FechaServidor,EntaEnLoc) VALUES ('" + strIDViaggio + "','" + strIDConsegna + "','999', Getdate() ,Getdate(),'0')"
                  If Not (DA.ExecuteQuery(SQL_GEN_EV)) Then
                  End If

End Function

Private Sub btn_pararev_Click()
Dim rowev As Variant
Dim lngev As Long
Dim x As Integer
Dim DA As CDataAccess
Set DA = DataConnect.DA
Dim sql_switch_parar_ev As String


If Not (DA.QueryCommand("SELECT SWITCH FROM PARAREVE", rowev, lngev)) Then
End If

x = (CInt(rowev(0, 0)) - 1) * (-1)
sql_switch_parar_ev = "UPDATE PARAREVE SET SWITCH =" + CStr(x)
If Not (DA.ExecuteQuery(sql_switch_parar_ev)) Then
End If


If (x = 0) Then
    MsgBox ("Parar Eventos está ahora DESACTIVADO")
Else
    MsgBox ("Parar Eventos está ahora ACTIVADO")
End If



End Sub
Private Function Captura_SMS() As Boolean
On Error GoTo ErrHandler
    Dim varLastID As Variant
    Dim varMinuti As Variant
    Dim DA As CDataAccess
    Set DA = DataConnect.DA
    Dim Varkbuffer As Variant
    Dim lngHowMany  As Long
    
    
    Dim KVE As CSatellitareEvento
    
    If Not DA.QueryCommand("SELECT ANUPROGR FROM ANAGANU0 WHERE (ANUIDANU = 'SMS')", Varkbuffer, lngHowMany) Then
    End If
    
    mskLastSMS = CLng(Varkbuffer(0, 0))
    
    varLastID = mskLastSMS.Text
    
    '--SM -JG Detección de ejecución de proccesos
    Dim ret As Boolean
    Dim x As CTextFile
    Set x = New CTextFile
    ret = x.OpenFile(MaskEdUnidad.Text & ":\Resoil2\Alarmas\SMS.txt", 1)
    ret = x.CloseFile()
    
    
    Set KVE = New CSatellitareEvento
    With KVE
        Set .DS.DA = DataConnect.DA
        Set .DS.ErrH = DataConnect.DA.ErrH
        
        KVE.Logger.LogFileName = gmstrServerLogFileName
        KVE.Logger.enableLevel (strLogLevels)
    
        XA.value(4, 1) = "En Ejecucion."
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
                    LogMessage "Inicio proceso captura SMS." & vbCrLf

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

    XA.value(4, 1) = "En Espera."
    gridProcedures.RefetchCol 1

    Captura_SMS = True
    Set KVE = Nothing
    GoTo EndProc

ErrHandler:
    Set KVE = Nothing
    
    mblnRunOk = True
    XA.value(0, 1) = "En Error."
    gridProcedures.RefetchCol 1
    
EndProc:
    If Not IsNull(varLastID) Then
    mskLastSMS.Text = varLastID
    DA.ExecuteQuery ("UPDATE ANAGANU0 SET ANUPROGR = '" & varLastID & "' WHERE (ANUIDANU = 'SMS')")
    gmRegistry.SaveSetting "Options", REGISTRY_KEY_NAME_LAST_SMS, varLastID, True
    End If
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
    Dim lngHowMany As Long
    Dim Varkbuffer As Variant
    
    Set XA = New XArrayDB

    XA.ReDim 0, 9, 0, 2
    
    XA.value(0, 0) = "Proceso de captura eventos Palms"
    XA.value(0, 1) = "Esperando"
    XA.value(0, 2) = "Ejecutar"
    
    XA.value(1, 0) = "Proceso de generación de GES TPH."
    XA.value(1, 1) = "Esperando"
    XA.value(1, 2) = "Ejecutar"
    
    XA.value(2, 0) = "Proceso de generación de GES EUT."
    XA.value(2, 1) = "Esperando"
    XA.value(2, 2) = "Ejecutar"
    
    XA.value(3, 0) = "Proceso de recepción de EcoGES."
    XA.value(3, 1) = "Esperando"
    XA.value(3, 2) = "Ejecutar"
    
    XA.value(4, 0) = "Proceso Captura de SMS"
    XA.value(4, 1) = "Esperando"
    XA.value(4, 2) = "Ejecutar"
    
    XA.value(5, 0) = "Proceso de recepción de Albaranes."
    XA.value(5, 1) = "Esperando"
    XA.value(5, 2) = "Ejecutar"

    XA.value(6, 0) = "Proceso de recepción precintos Oesia"
    XA.value(6, 1) = "Esperando"
    XA.value(6, 2) = "Ejecutar"

    XA.value(7, 0) = "Proceso de recepción precintos Thaumat"
    XA.value(7, 1) = "Esperando"
    XA.value(7, 2) = "Ejecutar"
    
    XA.value(8, 0) = "Proceso completado estados Compartimentos en GES"
    XA.value(8, 1) = "Esperando"
    XA.value(8, 2) = "Ejecutar"

    XA.value(9, 0) = "Proceso Generacion de desprecintados automática"
    XA.value(9, 1) = "Esperando"
    XA.value(9, 2) = "Ejecutar"
    

    
    mblnRunOk = True
     
    gridProcedures.Array = XA
    gridProcedures.ReBind
    
    MaskEdUnidad = gmRegistry.GetSetting("Options", REGISTRY_KEY_NAME_UNIDAD, "", True)
    
    Dim DA As CDataAccess
    Set DA = DataConnect.DA
    
  
    
    If Not DA.QueryCommand("SELECT ANUPROGR FROM ANAGANU0 WHERE (ANUIDANU = 'SMS')", Varkbuffer, lngHowMany) Then
    End If
    
    If Not (IsNull(CLng(Varkbuffer(0, 0)))) Then mskLastSMS = CLng(Varkbuffer(0, 0))
    
'    mskLastSMS = gmRegistry.GetSetting("Options", REGISTRY_KEY_NAME_LAST_SMS, "", True)
   
    SetTimerInterval CLng(gmRegistry.GetSetting("Options", REGISTRY_KEY_NAME_EXECUTE_SECONDS, "15", True)), True
    StartTimer

    strLogLevels = gmRegistry.GetSetting("Options", REGISTRY_KEY_NAME_LOG_LEVELS, "IDEW", True)
    Me.mskFileNameEUT = gmRegistry.GetSetting("Options", REGISTRY_KEY_NAME_PATH_EUT, "", True)
    Me.mskFileNameTPH = gmRegistry.GetSetting("Options", REGISTRY_KEY_NAME_PATH_TPH, "", True)
    Me.MaskEdHoras = gmRegistry.GetSetting("Options", REGISTRY_KEY_NAME_HORAS, "", True)
    Set mobjOutFile = New CTextFile
    
    
                With mobjOutFile
                .Add "Ente Emittente", 4, "1"
                .Add "Progressivo Invio", 10, "2", , fftNumerico
                .Add "Data Sistema", 8, "3"
                .Add "Ora Sistema", 8, "4"
                .Add "Numero Record", 4, "5", , fftNumerico
                .Add "Rif. Esterno 1 Consegna", 13, "6"
                .Add "Partizione", 2, "7", "00"
                .Add "Destinazione", 8, "8"
                .Add "GPS", 2, "9", "01", fftNumerico
                .Add "X GPS", 9, "10"
                .Add "Y GPS", 9, "11"
                .Add "Prodotto", 7, "12"
                .Add "Deposito", 4, "13"
                .Add "Q.tà Concordata Ordine", 9, "14", , fftNumerico
                .Add "Unità di Misura", 1, "15", "L"
                .Add "Data Viaggio", 8, "16"
                .Add "Data Consegna aggiornata", 8, "17"            ' Inizia a Colonna 107
                .Add "Ora Consegna aggiornata", 5, "18"
                .Add "Q.tà Trasportata", 9, "19", , fftNumerico
                .Add "Q.tà Scaricata", 9, "20", , fftNumerico
                '.Add "Data Viaggio 2", 8, "21"                       ACTUALIZACIÖN GES AGOSTO 2011 SM
                '.Add "Ora Fine Carico", 5, "22"
                '.Add "Data Consegna effettiva", 8, "23"             ' Inizia a Colonna 138
                '.Add "Ora Consegna effettiva", 5, "24"
                .Add "Badge UdT", 11, "21"
                .Add "Nr. Sequenza", 2, "22"
                .Add "Nr. Carta Identità", 20, "23"
                .Add "Km Definitivi", 4, "24", , fftNumerico
                .Add "Km CEPSA", 4, "25", , fftNumerico
                .Add "Quantità 1", 6, "26", "000.00", fftAlfa
                .Add "Quantità 2", 6, "27", "000.00", fftAlfa
                .Add "Quantità 3", 6, "28", "000.00", fftAlfa
                .Add "Numero Consegna", 2, "29", , fftNumerico
                .Add "Note Consegna", 40, "30"
                .Add "Codice evento", 4, "31"
                .Add "Data Evento", 8, "32"                         ' Inizia a Colonna 269
                .Add "Ora Evento", 5, "33"
                .Add "Stato compartimento 1", 4, "34"
                .Add "Stato compartimento 2", 4, "35"
                .Add "Stato compartimento 3", 4, "36"
                .Add "Stato compartimento 4", 4, "37"
                .Add "Stato compartimento 5", 4, "38"
                .Add "Stato compartimento 6", 4, "39"
                .Add "Stato compartimento 7", 4, "40"
                .Add "Filler", 55, "41"
                .Add "Flag consegna", 1, "42"
                .Add "Accordi con cliente", 1, "43"
                .Add "Ora carico prevista", 5, "44"
                .Add "Fecha Prevista de carga", 8, "45"
                .Add "Mat. Tractora", 18, "46"
                .Add "Mat. Cisterna", 18, "47"
                .Add "Mat 3", 18, "48" 'Matricula 3 Lo que no es tractora ni cisterna
                .Add "Motivo Evento(Inc)", 3, "49"
                .Add "Cantidad Involucrada", 9, "50"
                .Add "Prod 1", 7, "51"
                .Add "Prod 2", 7, "52"
                .Add "Resuelto medios propios", 1, "53"
                .Add "Tiempo paralizacion", 5, "54"
                .Add "Destinatario Desvio", 8, "55"
                .Add "Observacion Incidencia", 200, "56" 'Nuevo Campo incidencia JG 25/11/21
            End With

    Set mobjFileAlb = New CTextFile
    With mobjFileAlb
        .Add "TipoRegistro", 2, "1"
        .Add "CentroCarga", 4, "2"
        .Add "FechaSal", 8, "3"
        .Add "HoraSal", 4, "4"
        .Add "Albaran", 8, "5"
        .Add "AnyoAlbaran", 4, "6"
        .Add "VerAlbaran", 2, "7"
        .Add "PaisTractora", 2, "8"
        .Add "VehiculoCLH", 3, "9"
        .Add "MaticulaTractora", 10, "10"
        .Add "MatriculaCisterna", 10, "11"
        .Add "NumeroPedido", 13, "12"
        .Add "Particion", 2, "13"
        .Add "RefCargaCLH", 8, "14"
        .Add "CodProdCLH", 7, "15"
        .Add "CodigAditivo", 3, "16"
        .Add "CantidadPedido", 13, "17"
        .Add "Compartimento", 2, "18"
        .Add "TanqueSalida", 8, "19"
        .Add "CantidadCargada", 13, "20"
        .Add "UnidadMedido", 1, "21"
    End With
    
Set DataClass = New CViaggioSelezione
    With DataClass.DS
        Set .FormHandle = Me
        .FormType = ftEdtRcd1
        Set .DA = DataConnect.DA
        .Field("Ente Emittente").Required = True
        '.ConfigureInterface
        'ConfigureObjects Me, ColTags
        Set ErrH = New CErrorHandler
        Set .ErrH = ErrH
End With
    
    
    
End Sub
'Public Function GeneraGesTrujetamea() 'cheat!!!!!!!!!!
'Dim varLastID As Variant
'Dim varMinuti As Variant
'
'Dim KVE As CSatellitareEvento
'
'
'Set KVE = New CSatellitareEvento
'With KVE
'    Set .DS.DA = DataConnect.DA
'    Set .DS.ErrH = DataConnect.DA.ErrH
'
'    gridProcedures.RefetchCol 1
'
'
'
'                If Not KVE.ListaGES("2406089007") Then
'
'                End If
'
'
'End With
'
'LogMessage "Inicio poceso de Generación GES: " & vbCrLf
'gridProcedures.RefetchCol 1
'Set KVE = Nothing
'
'
'
'
'End Function

Public Function GeneraGes_TPH()
    Dim DA As CDataAccess
    Set DA = DataConnect.DA
    Dim i As Integer
    Dim Varkbuffer As Variant
    Dim lngHowMany As Long
    Dim SQL_GES_recuperar As String
    Dim NLote As String
    Dim FechaGL As String
    Dim HoraGL As String
    Dim NCampos As Integer
    On Error GoTo ErrHandler
    
    
    SQL_GES_recuperar = "SELECT TOP 100 * From VIAGGES0 WHERE (GESSEND0 = 0) AND GESCODTR = '624' AND GESTRCEP = '1' AND GESPARAR = 0 AND (GESNUMPD IS NOT NULL)"
    LogMessage "Inicio poceso de Generación GES de TPH: " & vbCrLf
  
    XA.value(1, 1) = "En Ejecucion"
    gridProcedures.RefetchCol 1
   
    If Not DA.QueryCommand(SQL_GES_recuperar, Varkbuffer, lngHowMany) Then
        'objLogger.logError "Query SQL_GES_recuperar."
        'GoTo ErrHandler
    End If
    
    If (lngHowMany > 0) Then
    
'            ' SELECIONAR EL NUMERO DE LOTE ("007"ID PARA TPH)

             With DataClass.DS.DA
                    NLote = .GetCounter("anaganu0", "anuprogr", "anuidanu", "007", 0)
                    If NLote = -1 Then
        '                ErrCode = InvalidCode
        '                GoTo ErrHandler
                    End If
                    If NLote > 9999999999# Then
                        NLote = 1
                    End If
                End With
'                NLote = Format$(Time, "hhmmss")
                
            FechaGL = Format$(Date, "YYYYMMDD")
            HoraGL = Format$(Hour(Now), "00") & ":" & Format$(Minute(Now), "00") & ":" & Format$(Second(Now), "00")
           Dim ficherotph As String
           
           If (Len(mskFileNameTPH.Text) < 2) Then
           ficherotph = MaskEdUnidad.Text & ":\cepsa\tph\salida"
           Else
           ficherotph = mskFileNameTPH.Text
           End If
          
          
          If Not mobjOutFile.OpenFile(ficherotph + "\GES_N" + NombreGes(NLote) + ".DAT", 1) Then
'            Dim nombrearch  As String
'            nombrearch = Format$(Date, "ddmmyy") & Format$(Time, "hhmmss")
'            If Not mobjOutFile.OpenFile(ficherotph + "\GES_TPH" + nombrearch + ".DAT", 1) Then
                'ErrCode = FileNotFound
                'GoTo ErrHandler
           End If
           
           
            NCampos = lngHowMany
                   
        'comprobamos que tenemos algo que enviar en el GES
        If lngHowMany > 0 Then
                
          For i = 0 To UBound(Varkbuffer, 2)
               With mobjOutFile
               .ClearFields
                    .Item("1").value = "0" + Varkbuffer(1, i)
                    'generados cuando se genera el GES
                    .Item("2").value = NLote 'numero de LOTE
                    .Item("3").value = FechaGL
                    .Item("4").value = HoraGL
                    .Item("5").value = NCampos
                    '----------
                    .Item("6").value = Varkbuffer(6, i)
                 'Numero de partición "00" si no hay.
                     If (Len(Varkbuffer(7, i)) < 2) Then
                    .Item("7").value = "00"
                     Else
                    .Item("7").value = Varkbuffer(7, i)
                     End If
                    .Item("8").value = Varkbuffer(8, i)
                    'Tipo de coordenada en nuestro caso siempre 01
                    .Item("9").value = "01"
                    .Item("10").value = Varkbuffer(10, i)
                    .Item("11").value = Varkbuffer(11, i)
                    .Item("12").value = Varkbuffer(12, i)
                    .Item("13").value = Varkbuffer(13, i)
                    .Item("14").value = Varkbuffer(14, i)
                    'Unidad de medida, ("L" o "K")
                    If (Varkbuffer(15, i) = "K") Then
                    .Item("15").value = Varkbuffer(15, i)
                    Else
                    .Item("15").value = "L"
                    End If
                    .Item("16").value = Varkbuffer(16, i)
                    
                   'hora actualizada
                    If (Varkbuffer(17, i) = "00:00") Then
                     .Item("17").value = Varkbuffer(23, i)
                    Else
                     .Item("17").value = Varkbuffer(17, i)
                    End If
                    If (Varkbuffer(18, i) = "19000101") Then
                    'fecha actualizada.
                     .Item("18").value = Varkbuffer(24, i)
                    Else
                     .Item("18").value = Varkbuffer(18, i)
                    End If
                    
                    
                    .Item("19").value = Varkbuffer(19, i)
                    .Item("20").value = Varkbuffer(20, i)
'                    .Item("21").value = Varkbuffer(21, i)
'                    .Item("22").value = Varkbuffer(22, i)
'                    .Item("23").value = Varkbuffer(23, i)
'                    .Item("24").value = Varkbuffer(24, i)
                    .Item("21").value = Varkbuffer(25, i)
                    .Item("22").value = Varkbuffer(26, i)
                    .Item("23").value = Varkbuffer(27, i)
                    .Item("24").value = Varkbuffer(28, i)
                    .Item("25").value = Varkbuffer(29, i)
                    If (IsNull(Varkbuffer(30, i))) Then
                    Else
                    .Item("26").value = Replace(Varkbuffer(30, i), ",", ".")
                    End If
                    
                    If (IsNull(Varkbuffer(31, i))) Then
                    Else
                    .Item("27").value = Replace(Varkbuffer(31, i), ",", ".")
                    End If
                     
                    If (IsNull(Varkbuffer(32, i))) Then
                    Else
                    .Item("28").value = Replace(Varkbuffer(32, i), ",", ".")
                    End If
                    
                    
                    
                    .Item("29").value = Varkbuffer(33, i)
                    .Item("30").value = Varkbuffer(34, i)
                    If (Trim(Varkbuffer(35, i)) = "999") Then
                     .Item("31").value = ""
                    Else
                     .Item("31").value = Varkbuffer(35, i)
                    End If
                    .Item("32").value = Varkbuffer(36, i)
                    .Item("33").value = Varkbuffer(37, i)
                    .Item("34").value = Varkbuffer(38, i)
                    .Item("35").value = Varkbuffer(39, i)
                    .Item("36").value = Varkbuffer(40, i)
                    .Item("37").value = Varkbuffer(41, i)
                    .Item("38").value = Varkbuffer(42, i)
                    .Item("39").value = Varkbuffer(43, i)
                    .Item("40").value = Varkbuffer(44, i)
                    .Item("41").value = Varkbuffer(45, i)
                    .Item("42").value = Varkbuffer(46, i)
                    .Item("43").value = 0 'varKBuffer(47, i) PACTAR CLIENTE
                    .Item("44").value = Varkbuffer(22, i) ' LA MISMA HORA QUE PARA EL TRANSPORTISTA
                    '.Item("49").value = Updatear a 1 'enviado
                    '.Item("50").value = Transmitir Cepsa = 1(si es 0 no enviar)
                    'NUEVOS CAMPOS
                    .Item("45").value = Varkbuffer(21, i) 'fecha prevista carga
                    Dim rowbufmat As Variant
                    Dim longmat As Long
                    
                    
                Dim sql_r As String
                          sql_r = "SELECT     ANAGAAM0_1.AAMMTCEP AS TRact, ANAGAAM0.AAMMTCEP AS CIst " & _
                         " FROM         ANAGAAM0 RIGHT OUTER JOIN " & _
                         " ANAGAUT0 ON ANAGAAM0.AAMIDAAM = ANAGAUT0.AUTIDAM2 LEFT OUTER JOIN " & _
                         " ANAGAAM0 AS ANAGAAM0_1 ON ANAGAUT0.AUTIDAM1 = ANAGAAM0_1.AAMIDAAM RIGHT OUTER JOIN " & _
                         " VIAGVVT0 ON ANAGAUT0.AUTIDAUT = VIAGVVT0.VVTIDAUT " & _
                         "WHERE     (VIAGVVT0.VVTIDVVT = ?) "




                    If Not DA.QueryCommand(sql_r, rowbufmat, longmat, CStr(Varkbuffer(54, i))) Then
                        'ERROR
                         GoTo ErrHandler
                    End If
                    If longmat > 0 Then
                    
                    .Item("46").value = rowbufmat(0, 0) 'mat trac
                    .Item("47").value = rowbufmat(1, 0) 'mat cis
                    End If
                    .Item("48").value = "" 'mat 3
                    .Item("49").value = Varkbuffer(56, i) ' Motivo INC
                    .Item("50").value = Varkbuffer(57, i) ' Cantidadsss
                    
                    Dim PROD As Variant
                    Dim longprod As Long
                    


                If Not IsNull(Varkbuffer(58, i)) And Varkbuffer(58, i) <> "" Then
                    If (Trim(Varkbuffer(35, i)) = "011") Then
                    .Item("51").value = Varkbuffer(58, i)
                    Else
                    .Item("51").value = ""
                    End If
                End If
                    
                If Not IsNull(Varkbuffer(59, i)) And Varkbuffer(59, i) <> "" Then
                    If (Trim(Varkbuffer(35, i)) = "011") Then
                    .Item("52").value = Varkbuffer(59, i)
                    Else
                    .Item("52").value = ""
                    End If
                End If
                
                
                If Varkbuffer(60, i) = True Then
                    .Item("53").value = "S" 'resuelto medios propios?
                Else
                       If (Trim(Varkbuffer(35, i)) = "012") Then
                             .Item("53").value = "N"
                        Else
                            .Item("53").value = ""
                        End If
                End If
                
                If Not IsNull(Varkbuffer(61, i)) Then
                        If (Trim(Varkbuffer(35, i)) = "019") Or (Trim(Varkbuffer(35, i)) = "020") Then
                             .Item("54").value = Replace(Format$(CStr(Varkbuffer(61, i)), "000.0"), ",", ".")
                        Else
                            .Item("54").value = ""
                        End If
                End If
                    
                ' Destinatario desvio
                If Not IsNull(Varkbuffer(62, i)) Then
                        If (Trim(Varkbuffer(35, i)) = "017") Then
                             .Item("55").value = Replace(Format$(CStr(Varkbuffer(62, i)), "000.0"), ",", ".")
                        Else
                            .Item("55").value = ""
                        End If
                End If
                    

        
                End With
              mobjOutFile.WriteLine
            Next i
        
           mobjOutFile.CloseFile
            
            'Updatear los capos generados en el GES
            Dim SQL_UPDATE_GESENVIADOS As String
            Dim SQL_UPDATE_GES As String
            SQL_UPDATE_GESENVIADOS = "UPDATE VIAGGES0 SET GESSEND0 = 1, GESNUMLO =? ,GESDAYDY = ?, GESDATHR = ? ,GESNUMRE = ?  WHERE GESIDENT = ?" 'LOTE,FECHA,HORA,NUMREGISTROS WHERE ID
        
            Dim IDGES As Long
            
            For i = 0 To UBound(Varkbuffer, 2)
                    IDGES = Varkbuffer(0, i)
                    If Not DA.ExecuteCommand(SQL_UPDATE_GESENVIADOS, NLote, FechaGL, HoraGL, NCampos, IDGES) Then
                    End If
            Next i
        End If
     End If
    
     

ErrHandler:

    XA.value(1, 1) = "Error"
    gridProcedures.RefetchCol 1

EndProc:

         XA.value(1, 1) = "En Espera"
        gridProcedures.RefetchCol 1
    

End Function
Private Function RecepcionEcoges()
Dim SQL_ECOGES As String
Dim DA As CDataAccess
Set DA = DataConnect.DA
Dim ArchOrigen, caracter, dirorigen, dirdestino As String
Dim cont, contlin, codcar, linfich As Integer
Dim contenido, car, linea, observa, lote, NumeroPD As String
Dim Can As Long
Dim PRCOD, EnteEm, numpr, Respuesta, CodEv As String
Dim Varkbuffer As Variant
Dim lngHowMany As Long
Dim i As Integer
Dim directorio(0 To 1, 0 To 1) As String
Dim Ficheros As String
Dim TotalFicheros As Long
Dim x As CTextFile
Set x = New CTextFile
Dim ret As Boolean
Dim Archivos As Integer
On Error GoTo ErrHandler

'SQL_ECOGES = "Update VIAGGES0 SET  GESRECEP = ?, GESOBCEP = ? WHERE (GESCODTR = ?) AND (GESNUMLO = ?) AND (GESPRCOD = ?) AND (GESCANPR = ?) AND (GESCODEV = ?) AND (GESNUMPR= ?)"
LogMessage "Inicio poceso de Recepción de EcoGES: " & vbCrLf
  
  XA.value(3, 1) = "Ejecutando"
  gridProcedures.RefetchCol 1

'Directorio Origen = (0,i) Directorio destino(1,i)
directorio(0, 0) = MaskEdUnidad.Text & ":\FICHEROS\TPH\Cepsa\Ecoges\"
directorio(1, 0) = MaskEdUnidad.Text & ":\FICHEROS\TPH\Cepsa\ecoges\leidos\"
directorio(0, 1) = MaskEdUnidad.Text & ":\FICHEROS\EUT\Cepsa\Ecoges\"
directorio(1, 1) = MaskEdUnidad.Text & ":\FICHEROS\EUT\Cepsa\ecoges\leidos\"


For i = 0 To 1
    ArchOrigen = Dir(directorio(0, i) + "ECO*.DAT")
    Archivos = 1
    Do While ArchOrigen <> ""
        Open (directorio(0, i) + ArchOrigen) For Input As #1
        contenido = Input(LOF(1), #1)
        linfich = 0
            For cont = 1 To Len(contenido)
                observa = ""
                linea = ""
                codcar = 0
                Do While (codcar <> 10 And codcar <> 13) And cont <= Len(contenido)
                    car = Mid(contenido, cont, 1)
                    linea = linea + car
                    cont = cont + 1
                    codcar = ASC(car)
                Loop
            'observa = Left(Mid(linea, 501, (Len(linea) - 441)), 90)
            ' JG 25/11/21 se cambia el ecoges para añadir observaciones incidencias
            observa = Left(Mid(linea, 804, (Len(linea) - 804)), 90)
            lote = Mid(linea, 8, 7)
            NumeroPD = Mid(linea, 35, 13)
            'EnteEm = Mid(linea, 2, 3)
            EnteEm = "004" ' Todo es Oilogist
            numpr = Mid(linea, 48, 2)
            Respuesta = Mid(linea, 500, 1)
            PRCOD = Mid(linea, 78, 7)
            'Can = Mid(linea, 89, 9)
            CodEv = Mid(linea, 239, 3)
            cont = cont - 1
            linfich = linfich + 1
            CodEv = Trim(CodEv)
            If (IsNull(CodEv) Or CodEv = "" Or IsEmpty(CodEv)) Then
                CodEv = "999"
            End If
            numpr = Trim(numpr)
            If (numpr = "00") Then
                numpr = ""
            End If
            
            ' Aqui puedes coger los campos de la linea que necesites para guardarlos en un variant más manejable
            SQL_ECOGES = "Update VIAGGES0 SET  GESRECEP ='" + Respuesta + "', GESOBCEP = '" + observa + "' WHERE (GESCODTR = '" + EnteEm + "') AND (GESNUMLO = '" + CStr(lote) + "') AND (GESPRCOD = '" + PRCOD + "') AND (GESNUMPD = '" + NumeroPD + "') AND (GESCODEV = '" + CodEv + "') AND (GESNUMPR= '" + numpr + "')"
            If Not DA.ExecuteQuery(SQL_ECOGES) Then
                 LogMessage "Error en la lectura del fichero: " + ArchOrigen + " Pedido: " + NumeroPD & vbCrLf
            End If
        
            Next cont
         Close #1
         
          FileCopy (directorio(0, i) + ArchOrigen), (directorio(1, i) + ArchOrigen)
          Kill (directorio(0, i) + ArchOrigen)
          ArchOrigen = Dir(directorio(0, i) + "ECO*.DAT")
          Archivos = Archivos + 1
          ' Limite de ficheros ecoges a tratar
          If Archivos > 40 Then ArchOrigen = ""
    Loop
Next i
         ' Reenvío de eventos Ges de los que no se ha recibido información
        Dim Horasreen As String
        Dim String_Reenvio As String
        If Not (IsNull(MaskEdHoras)) And Not (IsEmpty(MaskEdHoras)) Then
            Horasreen = MaskEdHoras / 24
            Horasreen = Replace(Horasreen, ",", ".")
            String_Reenvio = "UPDATE VIAGGES0 Set GESSEND0 = 0 WHERE (CONVERT(datetime, GESDAYDY + ' ' + GESDATHR, 112) > GETDATE() - 1) AND (GESRECEP IS NULL) AND (CONVERT(datetime,GESDAYDY + ' ' + GESDATHR, 112) < GETDATE() -" & Horasreen & ") AND (GESSEND0 <> 0)"
            If DA.ExecuteQuery(String_Reenvio) Then
            End If
        End If
' Miramos si funciona Fargo

Ficheros = Dir(MaskEdUnidad.Text & ":\FICHEROS\EUT\Cepsa\Salida\*.*")
TotalFicheros = 0
    While Ficheros <> ""
        TotalFicheros = TotalFicheros + 1
        Ficheros = Dir
    Wend
    
    If (TotalFicheros > 40) Then
        ret = x.OpenFile(MaskEdUnidad.Text & ":\Resoil2\Alarmas\fargo.txt", 1)
        ret = x.CloseFile()
    End If



ErrHandler:
    XA.value(3, 1) = "Error"
    gridProcedures.RefetchCol 1
    Close #1

EndProc:

     XA.value(3, 1) = "En Espera"
    gridProcedures.RefetchCol 1
    
End Function

Public Function NombreGes(NLote As String) As String
Dim i As Integer
Dim x As Integer
Dim R As String

x = Len(NLote)
R = CStr(NLote)
For i = x To 9
    R = "0" + R
Next i

NombreGes = R

End Function
Public Function Captura_albaranes()
Dim varBuffer As Variant
Dim SQL_UPDATE_ALB As String
Dim SQL_INSERT_ALB As String
Dim SQL_BUSCA_VIAJE As String
Dim DA As CDataAccess
Set DA = DataConnect.DA
Dim ArchOrigen, caracter, dirorigen, dirdestino As String
Dim cont, contlin, codcar, linfich As Integer
Dim contenido, car, linea, observa, lote, NumeroPD As String
Dim Can As Long
Dim PRCOD, EnteEm, numpr, Respuesta, CodEv As String
Dim Varkbuffer As Variant
Dim Varkbuffvi As Variant
Dim lngHowMany As Long
Dim lngHowvi As Long
Dim i, j As Integer
Dim directorio(0 To 1, 0 To 1) As String
Dim TipoRegistro As String
Dim CentroCarga As String
Dim FechaSal As String
Dim HoraSal As String
Dim Albaran As String
Dim AnyoAlbaran As String
Dim VerAlbaran As String
Dim PaisTractora As String
Dim VehiculoCLH As String
Dim MaticulaTractora As String
Dim MatriculaCisterna As String
Dim NumeroPedido As String
Dim Particion As String
Dim RefCargaCLH As String
Dim CodProdCLH As String
Dim CodigAditivo As String
Dim CantidadPedido As String
Dim Compartimento As String
Dim TanqueSalida As String
Dim CantidadCargada As String
Dim UnidadMedido As String
Dim NumViaje As String
Dim FechaSalida As String
Dim Fso As New FileSystemObject
Dim Archivo As File
Dim FechaCreacion As String
Dim Archivos As Integer
On Error GoTo ErrHandler

'SQL_ECOGES = "Update VIAGGES0 SET  GESRECEP = ?, GESOBCEP = ? WHERE (GESCODTR = ?) AND (GESNUMLO = ?) AND (GESPRCOD = ?) AND (GESCANPR = ?) AND (GESCODEV = ?) AND (GESNUMPR= ?)"
LogMessage "Inicio poceso Captura de albaranes: " & vbCrLf
  
  XA.value(5, 1) = "Ejecutando"
  gridProcedures.RefetchCol 1

'Directorio Origen = (0,i) Directorio destino(1,i)
directorio(0, 0) = MaskEdUnidad.Text & ":\FICHEROS\TPH\Cepsa\Albaranes\"
directorio(1, 0) = MaskEdUnidad.Text & ":\FICHEROS\TPH\Cepsa\Albaranes\leidos\"
directorio(0, 1) = MaskEdUnidad.Text & ":\FICHEROS\EUT\Cepsa\Albaranes\"
directorio(1, 1) = MaskEdUnidad.Text & ":\FICHEROS\EUT\Cepsa\Albaranes\leidos\"


For i = 0 To 1
    ArchOrigen = Dir(directorio(0, i) + "ALP*.*")
    Archivos = 0
    
    Do While ArchOrigen <> ""
    
        Set Archivo = Fso.GetFile(directorio(0, i) + ArchOrigen)
        FechaCreacion = Archivo.DateLastModified
        
    
        If Not mobjFileAlb.OpenFile(directorio(0, i) + ArchOrigen, 0) Then
            GoTo ErrHandler
        End If
        Do While Not mobjFileAlb.EOF
            varBuffer = mobjFileAlb.GetLine(False)
            If Not mobjFileAlb.ParseFixedLenFields(CStr(varBuffer), 1, 21) Then
                    GoTo ErrHandler
                    End If
            TipoRegistro = mobjFileAlb("1").value
            CentroCarga = mobjFileAlb("2").value
            FechaSal = mobjFileAlb("3").value
            HoraSal = mobjFileAlb("4").value
            Albaran = mobjFileAlb("5").value
            AnyoAlbaran = mobjFileAlb("6").value
            VerAlbaran = mobjFileAlb("7").value
            PaisTractora = mobjFileAlb("8").value
            VehiculoCLH = mobjFileAlb("9").value
            MaticulaTractora = mobjFileAlb("10").value
            MatriculaCisterna = mobjFileAlb("11").value
            NumeroPedido = mobjFileAlb("12").value
            Particion = mobjFileAlb("13").value
            RefCargaCLH = mobjFileAlb("14").value
            CodProdCLH = mobjFileAlb("15").value
            CodigAditivo = mobjFileAlb("16").value
            CantidadPedido = mobjFileAlb("17").value
            Compartimento = mobjFileAlb("18").value
            TanqueSalida = mobjFileAlb("19").value
            CantidadCargada = mobjFileAlb("20").value
            UnidadMedido = mobjFileAlb("21").value
            
            If Particion = "00" Then
                Particion = ""
            End If
            
            Dim añoc As String
            añoc = Mid(FechaSal, 1, 4)
            Dim mesc As String
            mesc = Mid(FechaSal, 5, 2)
            Dim diac As String
            diac = Mid(FechaSal, 7, 2)
            
            If Len(HoraSal) = 3 Then
            HoraSal = "0" & HoraSal
            End If
            
            HoraSal = Left(HoraSal, 2) & ":" & Right(HoraSal, 2)
            
           
            FechaSalida = diac & "/" & mesc & "/" & añoc & " " & HoraSal
                            
            
            If TipoRegistro = "41" Then
            
'            ' busco el viaje correspondiente
'
'            SQL_BUSCA_VIAJE = "SELECT VIAGVVC0.VVCIDVVT FROM VIAGVVC0 INNER JOIN ORDIOCT0 ON VIAGVVC0.VVCIDOCT = ORDIOCT0.OCTIDOCT WHERE (ORDIOCT0.OCTRIFE1 = ?)"
'
'
'                If DA.QueryCommand(SQL_BUSCA_VIAJE, Varkbuffvi, lngHowvi, Left(NumeroPedido, 10)) Then
'                End If
'                Numviaje = ""
'                    If lngHowvi > 0 Then
'                    Numviaje = (Varkbuffvi(0, 0))
'                    End If
'
                SQL_INSERT_ALB = "INSERT INTO ANAGAALCP (AALCDED0,AALCDTCA,AALCDTLL,AALCNRAL,AALCANAL,AALCVRAL,AALCPSAL,AALCVHCL,AALCMTTR, AALCMTCR,AALCNOTE,AALCNRPR,AALCRFCL,AALCIDPR,AALCIDAD,AALCCANP,AALCCOMP,AALCTQSA,AALCCANC,AALCUNMD) " & _
                                " VALUES ('" + CentroCarga + "','" + FechaSalida + "','" + FechaCreacion + "','" + Albaran + "','" + AnyoAlbaran + "','" + VerAlbaran + "','" + PaisTractora + "','" + VehiculoCLH + "','" + MaticulaTractora + "','" + MatriculaCisterna + "','" + NumeroPedido + "','" + Particion + "','" + RefCargaCLH + "','" + CodProdCLH + "','" + CodigAditivo + "','" + CantidadPedido + "','" + Compartimento + "','" + TanqueSalida + "','" + CantidadCargada + "','" + UnidadMedido + "')"
                If Not (DA.ExecuteQuery(SQL_INSERT_ALB)) Then
                     LogMessage "Error en Escritura de albaran: " & Albaran & vbCrLf
                End If
            End If
                    
           Loop


              mobjFileAlb.CloseFile
              FileCopy (directorio(0, i) + ArchOrigen), (directorio(1, i) + ArchOrigen)
              Kill (directorio(0, i) + ArchOrigen)
              ArchOrigen = Dir(directorio(0, i) + "ALP*.*")
              'Control archivos a procesar
              Archivos = Archivos + 1
              If Archivos > 10 Then ArchOrigen = ""
    
    Loop
Next i
 
 
ErrHandler:
    
    XA.value(5, 1) = "Error"
    gridProcedures.RefetchCol 1



EndProc:

     XA.value(5, 1) = "En Espera"
    gridProcedures.RefetchCol 1


End Function
Public Function Captura_prec_oesia()

Dim Varkbuffer As Variant
Dim VarkbufferCis As Variant
Dim lngHowMany As Long
Dim lngHowManyCis As Long
Dim SQL_CAPTURA_PREC As String
Dim SQL_INSERTA_PREC As String
Dim SQL_MARCA_PREC As String
Dim DA As CDataAccess
Dim RegistroLeid, MatCisterna, FechaEvento, Deposito, EstadoDep As String

Set DA = DataConnect.DA
Dim i As Integer

On Error GoTo ErrHandler


    LogMessage "Inicio poceso de Captura de estado precintos " & vbCrLf
    XA.value(6, 1) = "En Ejecucion."
    gridProcedures.RefetchCol 1
        



    SQL_CAPTURA_PREC = " SELECT     EventosCisternasMBflotas.IdRegistro, ANAGAAM0.AAMIDAAM, EventosCisternasMBflotas.FechaEvento, EventosCisternasMBflotas.Deposito, " & _
                        "                      ANAGAEP0.AEPESCEP, EventosCisternasMBflotas.Capturado " & _
                        "FROM         EventosCisternasMBflotas LEFT OUTER JOIN  ANAGAAM0 ON EventosCisternasMBflotas.CodigoVehiculo = ANAGAAM0.AAMIDPRE COLLATE Modern_Spanish_CI_AS LEFT OUTER JOIN " & _
                        "                      ANAGAEP0 ON EventosCisternasMBflotas.Subestructura2 = ANAGAEP0.AEPIDEPD COLLATE Modern_Spanish_CI_AS " & _
                        "WHERE     (EventosCisternasMBflotas.Capturado = 0) AND (ANAGAEP0.AEPESCEP IS NOT NULL) ORDER BY EventosCisternasMBflotas.IdRegistro "
       
    
    If Not DA.QueryCommand(SQL_CAPTURA_PREC, Varkbuffer, lngHowMany) Then
            LogMessage "Error en el proceso de captura de precintos. " & vbCrLf
            gridProcedures.RefetchCol 1
   End If

 For i = 0 To lngHowMany - 1
    
         RegistroLeid = (Varkbuffer(0, i))
         MatCisterna = (Varkbuffer(1, i))
         FechaEvento = (Varkbuffer(2, i))
         Deposito = Right(Varkbuffer(3, i), 1)
         EstadoDep = (Varkbuffer(4, i))
         
        ' SQL_INSERTA_PREC = "UPDATE ANAGAEC0 SET AECEPDP" + Deposito + " = " + EstadoDep + " , AECFEES" + Deposito + " = " + FechaEvento + " WHERE (AECIDAAM = '" + MatCisterna + "')"
        
' Comprobamos que exista la cisterna



    If Not DA.QueryCommand("SELECT AECIDAAM FROM ANAGAEC0 WHERE (AECIDAAM ='" & MatCisterna & "')", VarkbufferCis, lngHowManyCis) Then
    End If
    ' si no existe la insertamos
    If lngHowManyCis < 1 Then
        If Not (DA.ExecuteQuery("INSERT INTO ANAGAEC0 (AECIDAAM) VALUES('" & MatCisterna & "')")) Then
        End If
    End If
    
        
         SQL_INSERTA_PREC = "UPDATE ANAGAEC0 SET AECEPDP" & Deposito & "  =  " & EstadoDep & "  , AECFEES" & Deposito & "  = CONVERT(DATETIME, '" & FechaEvento & "', 103) WHERE (AECIDAAM = '" & MatCisterna & "')"
                          
                If Not (DA.ExecuteQuery(SQL_INSERTA_PREC)) Then
                             LogMessage "Error en Escritura de precinto: " & RegistroLeid & vbCrLf
                Else
                SQL_MARCA_PREC = "UPDATE    EventosCisternasMBflotas SET Capturado = 1 WHERE (IdRegistro =" & RegistroLeid & ")"
                    If Not (DA.ExecuteQuery(SQL_MARCA_PREC)) Then
                         LogMessage "Error en Escritura de precinto: " & RegistroLeid & vbCrLf
                    End If
                End If
        
 Next i
               
  
GoTo EndProc

ErrHandler:
    
    XA.value(6, 1) = "En Error."
    gridProcedures.RefetchCol 1
    Exit Function
    
EndProc:

    XA.value(6, 1) = "En Espera"
    gridProcedures.RefetchCol 1

End Function

Public Function Captura_desprecintados()

Dim Varkbuffer As Variant
Dim VarkbufferEven As Variant
Dim VarkbufferPed As Variant
Dim Varkbuffer008 As Variant
Dim VarkbufferEnt As Variant
Dim lngHowMany As Long
Dim lngHowManyEven As Long
Dim lngHowManyPed As Long
Dim lngHowMany008 As Long
Dim lngHowManyEnt As Long
Dim SQL_Captura_Desp As String
Dim SQL_Busca_Evento As String
Dim SQL_MARCA_PREC As String
Dim SQL_Pedido_Deposito As String
Dim SQL_Busca_008 As String
Dim SQL_Busca_Ent As String
Dim SQL_GEN_EV As String
Dim DifLat, DifLong As Double
Dim EstaenCli As Integer
Dim Radio As Integer

Dim DA As CDataAccess
Dim RegistroLeid, MatCisterna, FechaEvento, Deposito, Idcisterna, UltimoEvento, ViajeDesp, PedidoDesp, ParticionDesp, OrdenDesp, Latitud, Longitud As String

Dim LatitudCli, LongitudCli As String

Set DA = DataConnect.DA
Dim i, j As Integer

On Error GoTo ErrHandler

    LogMessage "Inicio poceso de Captura de desprecintados " & vbCrLf
    XA.value(9, 1) = "En Ejecucion."
    gridProcedures.RefetchCol 1
        

    ' Capturamos los primeros 100 desprecintados que tenemos pendientes en Eventoscisternas mbflotas

    SQL_Captura_Desp = "SELECT     TOP 100 IdRegistro, FechaEvento, CodigoVehiculo, Deposito, Latitud, Longitud FROM dbo.EventosCisternasMBFlotas " & _
            "WHERE     (Subestructura2 = 'DP') AND (Captprec = 0) AND (Subestructura3 = 'VF' OR Subestructura3 = 'AP' OR Subestructura3 = 'BH' OR Subestructura3 = 'XX') order by fechaevento"
    
    If Not DA.QueryCommand(SQL_Captura_Desp, Varkbuffer, lngHowMany) Then
            LogMessage "Error en el proceso de captura de desprecintados. " & vbCrLf
            gridProcedures.RefetchCol 1
   End If

 For i = 0 To lngHowMany - 1
 
        RegistroLeid = ""
        MatCisterna = ""
        FechaEvento = ""
        Deposito = ""
        Idcisterna = ""
        UltimoEvento = ""
        ViajeDesp = ""
        PedidoDesp = ""
        ParticionDesp = ""
        OrdenDesp = ""
        DifLat = 0
        DifLong = 0
        Latitud = ""
        Longitud = ""
        LatitudCli = ""
        LongitudCli = ""
        EstaenCli = 0
    
         RegistroLeid = (Varkbuffer(0, i))
         FechaEvento = (Varkbuffer(1, i))
         Idcisterna = (Varkbuffer(2, i))
         Deposito = (Varkbuffer(3, i))
         If Not (IsNull(Varkbuffer(4, i))) Then Latitud = (Varkbuffer(4, i))
         If Not (IsNull(Varkbuffer(5, i))) Then Longitud = (Varkbuffer(5, i))

         
    ' Buscamos en la tabla Viagges0 el último estado de la cisterna, si el último evento que tiene es un 004 quiere decir que está en viaje.
         
    SQL_Busca_Evento = " SELECT     TOP 1 dbo.VIAGGES0.GESCODEV, dbo.VIAGVVT0.VVTIDVVT, CONVERT(datetime, dbo.VIAGGES0.GESDYEV0 + ' ' + dbo.VIAGGES0.GESHREV0, 112) AS HORAEV " & _
                        "FROM         dbo.ANAGAAM0 INNER JOIN  dbo.VIAGVVT0 ON dbo.ANAGAAM0.AAMIDAAM = dbo.VIAGVVT0.VVTIDAM2 RIGHT OUTER JOIN  dbo.VIAGGES0 ON dbo.VIAGVVT0.VVTIDVVT = dbo.VIAGGES0.IDVVT " & _
                        "WHERE     (dbo.ANAGAAM0.AAMIDPRE = ?) AND (CONVERT(datetime, dbo.VIAGGES0.GESDYEV0 + ' ' + dbo.VIAGGES0.GESHREV0, 112) < CONVERT(DATETIME,'" & FechaEvento & "', 103)) AND (dbo.VIAGGES0.GESCODEV < '008') " & _
                        "ORDER BY CONVERT(datetime, dbo.VIAGGES0.GESDYEV0 + ' ' + dbo.VIAGGES0.GESHREV0, 112) DESC "


    If Not DA.QueryCommand(SQL_Busca_Evento, VarkbufferEven, lngHowManyEven, Idcisterna) Then
    End If
    
    If lngHowManyEven > 0 Then
    
    UltimoEvento = RTrim(VarkbufferEven(0, 0))
        
        If UltimoEvento = "004" Then
            
            ' Damos en principio por descontado que el viaje que está desprecintado es el que tiene el último 004 en ruta al cliente y buscamos en la tabla de albaranes
            ' por si lo tenemos yá, que pedido iba cargado en ese depósito.
            
            
            ViajeDesp = RTrim(VarkbufferEven(1, 0))
            
            SQL_Pedido_Deposito = "SELECT     dbo.ANAGAALCP.AALCNOTE, dbo.ANAGAALCP.AALCNRPR, dbo.ORDIOCT0.OCTIDOCT, dbo.ANAGADZ0.ADZCGPSY, dbo.ANAGADZ0.ADZCGPSX " & _
            "FROM         dbo.ORDIOCT0 INNER JOIN " & _
            "                      dbo.ORDIOCR0 ON dbo.ORDIOCT0.OCTIDOCT = dbo.ORDIOCR0.OCRIDOCT INNER JOIN " & _
            "                      dbo.VIAGVVP0 ON dbo.ORDIOCR0.OCRIDOCT = dbo.VIAGVVP0.VVPIDOCT AND dbo.ORDIOCR0.OCRIDOCR = dbo.VIAGVVP0.VVPIDOCR INNER JOIN " & _
            "                      dbo.VIAGVVC0 ON dbo.VIAGVVP0.VVPIDVVT = dbo.VIAGVVC0.VVCIDVVT AND dbo.VIAGVVP0.VVPIDVVC = dbo.VIAGVVC0.VVCIDVVC INNER JOIN " & _
            "                      dbo.ANAGADZ0 ON dbo.VIAGVVC0.VVCIDADZ = dbo.ANAGADZ0.ADZIDADZ RIGHT OUTER JOIN " & _
            "                      dbo.ANAGAALCP ON dbo.ORDIOCR0.OCRRIFE2 = RIGHT(dbo.ANAGAALCP.AALCNOTE, 3) AND " & _
            "                      dbo.ORDIOCT0.OCTRAUT4 = dbo.ANAGAALCP.AALCNRPR AND dbo.ORDIOCT0.OCTRIFE1 = LEFT(dbo.ANAGAALCP.AALCNOTE, 10) " & _
            "WHERE     (dbo.VIAGVVP0.VVPIDVVT = ?) AND (dbo.ANAGAALCP.AALCCOMP = ?)"
        
            If Not DA.QueryCommand(SQL_Pedido_Deposito, VarkbufferPed, lngHowManyPed, ViajeDesp, Deposito) Then
            End If
            
            ' Si encontramos el pedido insertamos un 008 en la tabla con ese pedido encontrado, buscando antes que no exista otro ya insertado, en cuyo caso lo saltamos.
            
            If lngHowManyPed > 0 Then
                PedidoDesp = RTrim(VarkbufferPed(0, 0))
                ParticionDesp = RTrim(VarkbufferPed(1, 0))
                OrdenDesp = RTrim(VarkbufferPed(2, 0))
                If Not (IsNull(VarkbufferPed(3, 0))) Then LatitudCli = VarkbufferPed(3, 0)
                If Not (IsNull(VarkbufferPed(4, 0))) Then LongitudCli = VarkbufferPed(4, 0)
                If Latitud <> "" And Longitud <> "" And LatitudCli <> "" And LongitudCli <> "" Then
                    DifLat = LatitudCli - Latitud
                    DifLong = LongitudCli - Longitud
                    Radio = 50
                    If (DifLat > 0 And DifLong > 0) And (DifLat < (0.00001 * Radio) And DifLong < (0.000001 * Radio)) Then EstaenCli = 1
                    If (DifLat < 0 And DifLong < 0) And (DifLat > -(0.00001 * Radio) And DifLong > -(0.000001 * Radio)) Then EstaenCli = 1
                    If (DifLat > 0 And DifLong < 0) And (DifLat < (0.00001 * Radio) And DifLong > -(0.000001 * Radio)) Then EstaenCli = 1
                    If (DifLat < 0 And DifLong > 0) And (DifLat > -(0.00001 * Radio) And DifLong < (0.000001 * Radio)) Then EstaenCli = 1
                    Latitud = Replace(Latitud, Chr(44), Chr(46))
                    Longitud = Replace(Longitud, Chr(44), Chr(46))
                End If
                SQL_Busca_008 = "SELECT IdRegistro, CodigoServicio, CodigoTarea FROM EventosMBFlotasmanual WHERE (CodigoEvento = '008') AND (CodigoServicio = ?) AND (CodigoTarea = ?)"
                If Not DA.QueryCommand(SQL_Busca_008, Varkbuffer008, lngHowMany008, ViajeDesp, OrdenDesp) Then
                End If
                       
                If lngHowMany008 = 0 Then
                        SQL_GEN_EV = "INSERT INTO EventosMBFlotasManual (IdTareaMB,IdUsuarioMB,CodigoServicio,CodigoTarea,CodigoEvento,FechaEvento,FechaServidor,EntaEnLoc,Latitud,Longitud) VALUES ('-33','999','" & ViajeDesp & "','" & OrdenDesp & "','008','" & FechaEvento & "','" & FechaEvento & "','" & EstaenCli & "','" & Latitud & "','" & Longitud & "')"
                        If DA.ExecuteQuery(SQL_GEN_EV) Then
                        End If
                End If
            
            ' si no encontramos nada en la tabla de albaranes, buscamos la siguiente netrega que tiene que realizar y mandamos un evento 008 buscando que no existo otro anterior.
            
            Else
                SQL_Busca_Ent = "SELECT     dbo.VIAGVVP0.VVPIDOCT, dbo.ANAGADZ0.ADZCGPSY, dbo.ANAGADZ0.ADZCGPSX " & _
                "FROM         dbo.VIAGVVP0 INNER JOIN " & _
                "                      dbo.VIAGVVC0 ON dbo.VIAGVVP0.VVPIDVVT = dbo.VIAGVVC0.VVCIDVVT AND dbo.VIAGVVP0.VVPIDVVC = dbo.VIAGVVC0.VVCIDVVC INNER JOIN " & _
                "                      dbo.ANAGADZ0 ON dbo.VIAGVVC0.VVCIDADZ = dbo.ANAGADZ0.ADZIDADZ " & _
                "WHERE     (dbo.VIAGVVC0.VVCFLST5 = 'S') AND (dbo.VIAGVVP0.VVPIDVVT = ?) "

                If Not DA.QueryCommand(SQL_Busca_Ent, VarkbufferEnt, lngHowManyEnt, ViajeDesp) Then
                End If
                If lngHowManyEnt > 0 Then
                    For j = 0 To lngHowManyEnt - 1
                    OrdenDesp = (VarkbufferEnt(0, j))
                If Not (IsNull(VarkbufferEnt(1, 0))) Then LatitudCli = VarkbufferEnt(1, 0)
                If Not (IsNull(VarkbufferEnt(2, 0))) Then LongitudCli = VarkbufferEnt(2, 0)
                If Latitud <> "" And Longitud <> "" And LatitudCli <> "" And LongitudCli <> "" Then
                    DifLat = LatitudCli - Latitud
                    DifLong = LongitudCli - Longitud
                    Radio = 50
                    If (DifLat > 0 And DifLong > 0) And (DifLat < (0.00001 * Radio) And DifLong < (0.000001 * Radio)) Then EstaenCli = 1
                    If (DifLat < 0 And DifLong < 0) And (DifLat > -(0.00001 * Radio) And DifLong > -(0.000001 * Radio)) Then EstaenCli = 1
                    If (DifLat > 0 And DifLong < 0) And (DifLat < (0.00001 * Radio) And DifLong > -(0.000001 * Radio)) Then EstaenCli = 1
                    If (DifLat < 0 And DifLong > 0) And (DifLat > -(0.00001 * Radio) And DifLong < (0.000001 * Radio)) Then EstaenCli = 1
                    Latitud = Replace(Latitud, Chr(44), Chr(46))
                    Longitud = Replace(Longitud, Chr(44), Chr(46))
                End If
                        SQL_Busca_008 = "SELECT IdRegistro, CodigoServicio, CodigoTarea FROM EventosMBFlotasmanual WHERE (CodigoEvento = '008') AND (CodigoServicio = ?) AND (CodigoTarea = ?)"
                        If Not DA.QueryCommand(SQL_Busca_008, Varkbuffer008, lngHowMany008, ViajeDesp, OrdenDesp) Then
                        End If
                                   
                            If lngHowMany008 = 0 Then
                                    SQL_GEN_EV = "INSERT INTO EventosMBFlotasManual (IdTareaMB,IdUsuarioMB,CodigoServicio,CodigoTarea,CodigoEvento,FechaEvento,FechaServidor,EntaEnLoc,Latitud,Longitud) VALUES ('-33','999','" & ViajeDesp & "','" & OrdenDesp & "','008','" & FechaEvento & "','" & FechaEvento & "','" & EstaenCli & "','" & Latitud & "','" & Longitud & "')"
                                    If DA.ExecuteQuery(SQL_GEN_EV) Then
                                    End If
                            End If
                    Next j
                End If
            End If
        Else
        
       
        End If
        SQL_MARCA_PREC = "UPDATE    EventosCisternasMBflotas SET Captprec = 1 WHERE (IdRegistro =" & RegistroLeid & ")"
        If Not (DA.ExecuteQuery(SQL_MARCA_PREC)) Then
           LogMessage "Error en Escritura de desprecintado: " & RegistroLeid & vbCrLf
        End If
 
    Else
    SQL_MARCA_PREC = "UPDATE    EventosCisternasMBflotas SET Captprec = 1 WHERE (IdRegistro =" & RegistroLeid & ")"
        If Not (DA.ExecuteQuery(SQL_MARCA_PREC)) Then
           LogMessage "Error en Escritura de desprecintado: " & RegistroLeid & vbCrLf
        End If
 

    End If
         
 Next i
               
  
GoTo EndProc

ErrHandler:
    
    XA.value(9, 1) = "En Error."
    gridProcedures.RefetchCol 1
    Exit Function
    
EndProc:

    LogMessage "Fin poceso de Captura de desprecintados " & vbCrLf
    XA.value(9, 1) = "En Espera"
    gridProcedures.RefetchCol 1

End Function
Public Function Captura_prec_thaumat()

Dim Varkbuffer As Variant
Dim VarkbufferEC0 As Variant
Dim VarkbufferEDep As Variant
Dim VarkbufferPOS As Variant
Dim lngHowMany As Long
Dim lngHowManyEC0 As Long
Dim lngHowManyEDep As Long
Dim lngHowManyPOS As Long
Dim DA As CDataAccess
Dim RegistroLeid, MatCisterna, MatCisternaTh, FechaEvento, Deposito, EstadoDep, Latitudgr, Longitudgr, Mattract, CodGPS, FechaPos As String
Dim THDepo1, THDepo2, THDepo3, THDepo4, THDepo5, THDepo6, ESDepo1, ESDepo2, ESDepo3, ESDepo4, ESDepo5, ESDepo6, EDepCep, EDepRes As String
Dim RESDepo1, RESDepo2, RESDepo3, RESDepo4, RESDepo5, RESDepo6 As String
Dim SQL_ULT_EST_PREC As String
Dim SQL_INSERT_OECIS As String
Dim SQL_MARCA_PREC As String
Dim SQL_INSERTA_POS As String
Dim SQL_BUSCA_TRACT As String
Set DA = DataConnect.DA
Dim i As Integer
Dim h As Integer
Dim Latitud, Longitud As String
Dim Signolat, Signolong, LongitudCampo As Integer
Dim SQL_INSERTA_PREC As String
Dim Speed As Integer
On Error GoTo ErrHandler


LogMessage "Inicio poceso de Captura precintos Thaumat" & vbCrLf
XA.value(7, 1) = "En Ejecucion."
gridProcedures.RefetchCol 1


SQL_ULT_EST_PREC = "SELECT     dbo.ANAGAAM0.AAMIDPRE, dbo.ANAGAEP0.AEPESTHM AS DP1, ANAGAEP0_1.AEPESTHM AS DP2, ANAGAEP0_2.AEPESTHM AS DP3, " & _
"                      ANAGAEP0_3.AEPESTHM AS DP4, ANAGAEP0_4.AEPESTHM AS DP5, ANAGAEP0_5.AEPESTHM AS DP6, " & _
"dbo.ANAGAEC0.AECEPDP1,dbo.ANAGAEC0.AECEPDP2, dbo.ANAGAEC0.AECEPDP3, dbo.ANAGAEC0.AECEPDP4, dbo.ANAGAEC0.AECEPDP5, dbo.ANAGAEC0.AECEPDP6 " & _
"FROM         dbo.ANAGAEC0 LEFT OUTER JOIN  dbo.ANAGAEP0 ON dbo.ANAGAEC0.AECEPDP1 = dbo.ANAGAEP0.AEPESCEP LEFT OUTER JOIN " & _
"                      dbo.ANAGAEP0 ANAGAEP0_5 ON dbo.ANAGAEC0.AECEPDP6 = ANAGAEP0_5.AEPESCEP LEFT OUTER JOIN " & _
"                      dbo.ANAGAEP0 ANAGAEP0_4 ON dbo.ANAGAEC0.AECEPDP5 = ANAGAEP0_4.AEPESCEP LEFT OUTER JOIN " & _
"                      dbo.ANAGAEP0 ANAGAEP0_3 ON dbo.ANAGAEC0.AECEPDP4 = ANAGAEP0_3.AEPESCEP LEFT OUTER JOIN " & _
"                      dbo.ANAGAEP0 ANAGAEP0_2 ON dbo.ANAGAEC0.AECEPDP3 = ANAGAEP0_2.AEPESCEP LEFT OUTER JOIN " & _
"                      dbo.ANAGAEP0 ANAGAEP0_1 ON dbo.ANAGAEC0.AECEPDP2 = ANAGAEP0_1.AEPESCEP LEFT OUTER JOIN " & _
"                      dbo.ANAGAAM0 ON dbo.ANAGAEC0.AECIDAAM = dbo.ANAGAAM0.AAMIDAAM WHERE     (dbo.ANAGAAM0.AAMIDPRE = ?)"
        
' leemos los precintos por capturar
If Not DA.QueryCommand("SELECT * FROM EVEN_THMAT_RES ORDER BY DATETIME ASC", Varkbuffer, lngHowMany) Then
            LogMessage "Error en el proceso de captura de precintos Thaumat. " & vbCrLf
            gridProcedures.RefetchCol 1
End If

For i = 0 To lngHowMany - 1
    FechaPos = Varkbuffer(3, i)
    RegistroLeid = Varkbuffer(0, i)
    MatCisternaTh = Varkbuffer(1, i)
    If Not IsNull(Varkbuffer(2, i)) Then
        MatCisterna = Varkbuffer(2, i)
        Else
        MatCisterna = ""
        End If
    If Not (IsNull(Varkbuffer(5, i))) Then Latitudgr = Varkbuffer(5, i)
    If Not (IsNull(Varkbuffer(6, i))) Then Longitudgr = Varkbuffer(6, i)
    If Varkbuffer(7, i) = "S" Then Signolat = -1 Else Signolat = 1
    If Varkbuffer(8, i) = "W" Then Signolong = -1 Else Signolong = 1
    If Latitudgr <> "" And Longitudgr <> "" Then
            If Len(Longitudgr) < 7 Then
                LongitudCampo = Len(Longitudgr)
                Do While LongitudCampo < 7
                Longitudgr = "0" & Longitudgr
                LongitudCampo = LongitudCampo + 1
                Loop
            End If
    
    Latitud = CInt(Left(Latitudgr, 2)) + (CDbl((Mid(Latitudgr, 3, 2)) & "," & Right(Latitudgr, 4)) / 60)
    Longitud = CInt(Left(Longitudgr, 1)) + (CDbl((Mid(Longitudgr, 2, 2)) & "," & Right(Longitudgr, 4)) / 60)
    Latitud = Latitud * Signolat
    Longitud = Longitud * Signolong
    Latitud = Replace(Latitud, Chr(44), Chr(46))
    Longitud = Replace(Longitud, Chr(44), Chr(46))
    Speed = 0
    If Not (IsNull(Varkbuffer(9, i))) Then Speed = CInt(Varkbuffer(9, i))
    If Speed = 255 Then Speed = 0
    Speed = Speed * 1.852
    
    End If
    ' buscamos cual era el ultimo estado de la cisterna en nuestra bd
    If Not DA.QueryCommand(SQL_ULT_EST_PREC, VarkbufferEC0, lngHowManyEC0, MatCisternaTh) Then
            LogMessage "Error en el proceso de captura de precintos Thaumat. " & vbCrLf
            gridProcedures.RefetchCol 1
    End If
    ' si la cisterna no estaba en nuestro histórico la damos de alta y volvemos a mirar los estados aunque esten a vacio para seguir
    If lngHowManyEC0 = 0 And MatCisterna <> "" Then
        If Not (DA.ExecuteQuery("INSERT INTO ANAGAEC0 (AECIDAAM) VALUES('" & MatCisterna & "')")) Then
        End If
        If Not DA.QueryCommand(SQL_ULT_EST_PREC, VarkbufferEC0, lngHowManyEC0, MatCisternaTh) Then
                LogMessage "Error en el proceso de captura de precintos Thaumat. " & vbCrLf
                gridProcedures.RefetchCol 1
        End If
    End If
    If lngHowManyEC0 > 0 Then
    FechaEvento = Varkbuffer(3, i)
    THDepo1 = Right(RTrim(Varkbuffer(21, i)), 2)
    THDepo2 = Right(RTrim(Varkbuffer(22, i)), 2)
    THDepo3 = Right(RTrim(Varkbuffer(23, i)), 2)
    THDepo4 = Right(RTrim(Varkbuffer(24, i)), 2)
    THDepo5 = Right(RTrim(Varkbuffer(25, i)), 2)
    THDepo6 = Right(RTrim(Varkbuffer(26, i)), 2)
    If Not (IsNull(VarkbufferEC0(1, 0))) Then ESDepo1 = VarkbufferEC0(1, 0)
    If Not (IsNull(VarkbufferEC0(2, 0))) Then ESDepo2 = VarkbufferEC0(2, 0)
    If Not (IsNull(VarkbufferEC0(3, 0))) Then ESDepo3 = VarkbufferEC0(3, 0)
    If Not (IsNull(VarkbufferEC0(4, 0))) Then ESDepo4 = VarkbufferEC0(4, 0)
    If Not (IsNull(VarkbufferEC0(5, 0))) Then ESDepo5 = VarkbufferEC0(5, 0)
    If Not (IsNull(VarkbufferEC0(6, 0))) Then ESDepo6 = VarkbufferEC0(6, 0)
    If Not (IsNull(VarkbufferEC0(7, 0))) Then RESDepo1 = VarkbufferEC0(7, 0)
    If Not (IsNull(VarkbufferEC0(8, 0))) Then RESDepo2 = VarkbufferEC0(8, 0)
    If Not (IsNull(VarkbufferEC0(9, 0))) Then RESDepo3 = VarkbufferEC0(9, 0)
    If Not (IsNull(VarkbufferEC0(10, 0))) Then RESDepo4 = VarkbufferEC0(10, 0)
    If Not (IsNull(VarkbufferEC0(11, 0))) Then RESDepo5 = VarkbufferEC0(11, 0)
    If Not (IsNull(VarkbufferEC0(12, 0))) Then RESDepo6 = VarkbufferEC0(12, 0)
    
    ' Comparamos deposito a depósito los estados y si han cambiado los insetamos
    
        If THDepo1 <> ESDepo1 Then
        
            ' Capturamos el código Cepsa del estado a guardar
            EDepCep = ""
            EDepRes = ""
            If Not DA.QueryCommand("SELECT AEPIDEPD, AEPESCEP FROM ANAGAEP0 WHERE (AEPESTHM = ?)", VarkbufferEDep, lngHowManyEDep, THDepo1) Then
                    LogMessage "Error en el proceso de captura de precintos Thaumat. " & vbCrLf
                    gridProcedures.RefetchCol 1
            End If
                If lngHowManyEDep > 0 Then
                    EDepCep = VarkbufferEDep(0, 0)
                    EDepRes = VarkbufferEDep(1, 0)
                    SQL_INSERT_OECIS = "INSERT INTO EventosCisternasMBFlotas (CodigoVehiculo, FechaEvento, FechaServidor, TipoEvento, TotalEventos, NumeroEvento, Deposito, Subestructura1, Subestructura2, Subestructura3,Latitud, Longitud, Capturado, Velocidad) " & _
                    "VALUES    ('" & MatCisternaTh & "', '" & FechaEvento & "', GETDATE(), 'T', '1', '1', '01', 'TR', '" & EDepCep & "', 'XX','" & Latitud & "','" & Longitud & "','1','" & Speed & "')"
                    If Not (DA.ExecuteQuery(SQL_INSERT_OECIS)) Then
                    End If
                    
                    
                    SQL_INSERTA_PREC = "UPDATE ANAGAEC0 SET AECEPDP1 =  " & EDepRes & "  , AECFEES1  = CONVERT(DATETIME, '" & FechaEvento & "', 103) WHERE (AECIDAAM = '" & MatCisterna & "')"
                              
                    If Not (DA.ExecuteQuery(SQL_INSERTA_PREC)) Then
                    End If
                End If
        
        End If
        If THDepo2 <> ESDepo2 Then
        
            ' Capturamos el código Cepsa del estado a guardar
            EDepCep = ""
            EDepRes = ""
            If Not DA.QueryCommand("SELECT AEPIDEPD,AEPESCEP FROM ANAGAEP0 WHERE (AEPESTHM = ?)", VarkbufferEDep, lngHowManyEDep, THDepo2) Then
                    LogMessage "Error en el proceso de captura de precintos Thaumat. " & vbCrLf
                    gridProcedures.RefetchCol 1
            End If
                If lngHowManyEDep > 0 Then
                    EDepCep = VarkbufferEDep(0, 0)
                    EDepRes = VarkbufferEDep(1, 0)
                    SQL_INSERT_OECIS = "INSERT INTO EventosCisternasMBFlotas (CodigoVehiculo, FechaEvento, FechaServidor, TipoEvento, TotalEventos, NumeroEvento, Deposito, Subestructura1, Subestructura2, Subestructura3,Latitud, Longitud, Capturado, Velocidad) " & _
                    "VALUES    ('" & MatCisternaTh & "', '" & FechaEvento & "', GETDATE(), 'T', '1', '1', '02', 'TR', '" & EDepCep & "', 'XX','" & Latitud & "','" & Longitud & "','1'," & Speed & ")"
                    If Not (DA.ExecuteQuery(SQL_INSERT_OECIS)) Then
    
                    End If
                    SQL_INSERTA_PREC = "UPDATE ANAGAEC0 SET AECEPDP2 =  " & EDepRes & "  , AECFEES2  = CONVERT(DATETIME, '" & FechaEvento & "', 103) WHERE (AECIDAAM = '" & MatCisterna & "')"
                              
                    If Not (DA.ExecuteQuery(SQL_INSERTA_PREC)) Then
                    End If
                End If
        
        
        End If
        If THDepo3 <> ESDepo3 Then
        
            ' Capturamos el código Cepsa del estado a guardar
            EDepCep = ""
            EDepRes = ""
            If Not DA.QueryCommand("SELECT AEPIDEPD,AEPESCEP FROM ANAGAEP0 WHERE (AEPESTHM = ?)", VarkbufferEDep, lngHowManyEDep, THDepo3) Then
                    LogMessage "Error en el proceso de captura de precintos Thaumat. " & vbCrLf
                    gridProcedures.RefetchCol 1
            End If
                If lngHowManyEDep > 0 Then
                    EDepCep = VarkbufferEDep(0, 0)
                    EDepRes = VarkbufferEDep(1, 0)
                    SQL_INSERT_OECIS = "INSERT INTO EventosCisternasMBFlotas (CodigoVehiculo, FechaEvento, FechaServidor, TipoEvento, TotalEventos, NumeroEvento, Deposito, Subestructura1, Subestructura2, Subestructura3,Latitud, Longitud, Capturado, Velocidad) " & _
                    "VALUES    ('" & MatCisternaTh & "', '" & FechaEvento & "', GETDATE(), 'T', '1', '1', '03', 'TR', '" & EDepCep & "', 'XX','" & Latitud & "','" & Longitud & "','1'," & Speed & ")"
                    If Not (DA.ExecuteQuery(SQL_INSERT_OECIS)) Then
                    End If
                     SQL_INSERTA_PREC = "UPDATE ANAGAEC0 SET AECEPDP3 =  " & EDepRes & "  , AECFEES3  = CONVERT(DATETIME, '" & FechaEvento & "', 103) WHERE (AECIDAAM = '" & MatCisterna & "')"
                              
                    If Not (DA.ExecuteQuery(SQL_INSERTA_PREC)) Then
                    End If
                End If
        
        
        End If
        If THDepo4 <> ESDepo4 Then
        
            ' Capturamos el código Cepsa del estado a guardar
            EDepCep = ""
            EDepRes = ""
            If Not DA.QueryCommand("SELECT AEPIDEPD,AEPESCEP FROM ANAGAEP0 WHERE (AEPESTHM = ?)", VarkbufferEDep, lngHowManyEDep, THDepo4) Then
                    LogMessage "Error en el proceso de captura de precintos Thaumat. " & vbCrLf
                    gridProcedures.RefetchCol 1
            End If
                If lngHowManyEDep > 0 Then
                    EDepCep = VarkbufferEDep(0, 0)
                    EDepRes = VarkbufferEDep(1, 0)
                    SQL_INSERT_OECIS = "INSERT INTO EventosCisternasMBFlotas (CodigoVehiculo, FechaEvento, FechaServidor, TipoEvento, TotalEventos, NumeroEvento, Deposito, Subestructura1, Subestructura2, Subestructura3,Latitud, Longitud, Capturado, Velocidad) " & _
                    "VALUES    ('" & MatCisternaTh & "', '" & FechaEvento & "', GETDATE(), 'T', '1', '1', '04', 'TR', '" & EDepCep & "', 'XX','" & Latitud & "','" & Longitud & "','1'," & Speed & ")"
                    If Not (DA.ExecuteQuery(SQL_INSERT_OECIS)) Then
                    End If
    
                     SQL_INSERTA_PREC = "UPDATE ANAGAEC0 SET AECEPDP4 =  " & EDepRes & "  , AECFEES4  = CONVERT(DATETIME, '" & FechaEvento & "', 103) WHERE (AECIDAAM = '" & MatCisterna & "')"
                              
                    If Not (DA.ExecuteQuery(SQL_INSERTA_PREC)) Then
                    End If
                End If
        
        End If
        If THDepo5 <> ESDepo5 Then
        
            ' Capturamos el código Cepsa del estado a guardar
            EDepCep = ""
            EDepRes = ""
            If Not DA.QueryCommand("SELECT AEPIDEPD,AEPESCEP FROM ANAGAEP0 WHERE (AEPESTHM = ?)", VarkbufferEDep, lngHowManyEDep, THDepo5) Then
                    LogMessage "Error en el proceso de captura de precintos Thaumat. " & vbCrLf
                    gridProcedures.RefetchCol 1
            End If
                If lngHowManyEDep > 0 Then
                    EDepCep = VarkbufferEDep(0, 0)
                    EDepRes = VarkbufferEDep(1, 0)
                    SQL_INSERT_OECIS = "INSERT INTO EventosCisternasMBFlotas (CodigoVehiculo, FechaEvento, FechaServidor, TipoEvento, TotalEventos, NumeroEvento, Deposito, Subestructura1, Subestructura2, Subestructura3,Latitud, Longitud, Capturado, Velocidad) " & _
                    "VALUES    ('" & MatCisternaTh & "', '" & FechaEvento & "', GETDATE(), 'T', '1', '1', '05', 'TR', '" & EDepCep & "', 'XX','" & Latitud & "','" & Longitud & "','1'," & Speed & ")"
                    If Not (DA.ExecuteQuery(SQL_INSERT_OECIS)) Then
                    End If
    
                    SQL_INSERTA_PREC = "UPDATE ANAGAEC0 SET AECEPDP5 =  " & EDepRes & "  , AECFEES5  = CONVERT(DATETIME, '" & FechaEvento & "', 103) WHERE (AECIDAAM = '" & MatCisterna & "')"
                              
                    If Not (DA.ExecuteQuery(SQL_INSERTA_PREC)) Then
                    End If
                End If
        
        End If
        
        If THDepo6 <> ESDepo6 Then
        
            ' Capturamos el código Cepsa del estado a guardar
            EDepCep = ""
            EDepRes = ""
            If Not DA.QueryCommand("SELECT AEPIDEPD,AEPESCEP FROM ANAGAEP0 WHERE (AEPESTHM = ?)", VarkbufferEDep, lngHowManyEDep, THDepo6) Then
                    LogMessage "Error en el proceso de captura de precintos Thaumat. " & vbCrLf
                    gridProcedures.RefetchCol 1
            End If
                If lngHowManyEDep > 0 Then
                    EDepCep = VarkbufferEDep(0, 0)
                    EDepRes = VarkbufferEDep(1, 0)
                    SQL_INSERT_OECIS = "INSERT INTO EventosCisternasMBFlotas (CodigoVehiculo, FechaEvento, FechaServidor, TipoEvento, TotalEventos, NumeroEvento, Deposito, Subestructura1, Subestructura2, Subestructura3,Latitud, Longitud, Capturado, Velocidad) " & _
                    "VALUES    ('" & MatCisternaTh & "', '" & FechaEvento & "', GETDATE(), 'T', '1', '1', '06', 'TR', '" & EDepCep & "', 'XX','" & Latitud & "','" & Longitud & "','1'," & Speed & ")"
                    If Not (DA.ExecuteQuery(SQL_INSERT_OECIS)) Then
                    End If
    
                    SQL_INSERTA_PREC = "UPDATE ANAGAEC0 SET AECEPDP6 =  " & EDepRes & "  , AECFEES6  = CONVERT(DATETIME, '" & FechaEvento & "', 103) WHERE (AECIDAAM = '" & MatCisterna & "')"
                              
                    If Not (DA.ExecuteQuery(SQL_INSERTA_PREC)) Then
                    End If
                End If
        
        End If

    End If
        ' Insertamos la posición en la tabla PosicionesOesia
        
        ' Capturamos todas las tractoras que puedan enganchar esta cisterna
        
        
'        Latitud = Left(Latitud, 8)
'        Longitud = Left(Longitud, 8)
'
'            SQL_BUSCA_TRACT = "SELECT     ANAGAAM0.AAMIDDEU AS GPS, VIAGVVT0.VVTIDAM2 " & _
'                                "FROM         VIAGVVT0 INNER JOIN ANAGAAM0 ON VIAGVVT0.VVTIDAM1 = ANAGAAM0.AAMIDAAM " & _
'                                "WHERE     (VIAGVVT0.VVTDTINC > GETDATE() - 3) AND (VIAGVVT0.VVTIDASC = '038') " & _
'                                "GROUP BY ANAGAAM0.AAMIDDEU, VIAGVVT0.VVTIDAM2 " & _
'                                "HAVING  (ANAGAAM0.AAMIDDEU >= '800') AND (VIAGVVT0.VVTIDAM2 = ?)   "
'
'
'            If Not DA.QueryCommand(SQL_BUSCA_TRACT, VarkbufferPOS, lngHowManyPOS, MatCisterna) Then
'                        LogMessage "Error en el proceso de captura de tractoras " & vbCrLf
'                        gridProcedures.RefetchCol 1
'            End If
'
'        For h = 0 To lngHowManyPOS - 1
'            CodGPS = VarkbufferPOS(0, h)
'            Latitud = Replace(Latitud, ".", ",")
'            Longitud = Replace(Longitud, ".", ",")
'            SQL_INSERTA_POS = "Insert Into PosicionesOesia (IDVEHIC, FEUTCVAL, FEUTCULT, FESERV, FEACT, Longitud, Latitud) Values (" & CodGPS & ",'" & FechaPos & " ','" & FechaPos & "',Getdate(),'" & FechaPos & "','" & Latitud & "','" & Longitud & "')"
'            If Not (DA.ExecuteQuery(SQL_INSERTA_POS)) Then
'               LogMessage "Error en Escritura de posicion: " & RegistroLeid & vbCrLf
'            End If
'        Next h
        ' Marcamos el registro como capturado
        
        SQL_MARCA_PREC = "UPDATE    EVENTOS_THMAT SET Capturado = 1 WHERE (CODPOSITION =" & RegistroLeid & ")"
        If Not (DA.ExecuteQuery(SQL_MARCA_PREC)) Then
           LogMessage "Error en Escritura de desprecintado: " & RegistroLeid & vbCrLf
        End If

Next i
        ' Borrar tabla position
'        If Not (DA.ExecuteQuery("DELETE FROM EVENTOS_THMAT WHERE (DATETIME < GETDATE() - 5)")) Then
'           LogMessage "Error borrado eventos: " & RegistroLeid & vbCrLf
'        End If
'
   
GoTo EndProc

ErrHandler:
    
    XA.value(7, 1) = "En Error."
    gridProcedures.RefetchCol 1
    Exit Function
    
EndProc:
    LogMessage "Fin poceso de Captura precintos Thaumat" & vbCrLf
    XA.value(7, 1) = "En Espera"
    gridProcedures.RefetchCol 1

End Function
Public Function Rellena_Compartimentos()
Dim Varkbuffer As Variant
Dim VarkbufferECis As Variant
Dim VarkbufferEDep As Variant
Dim Varkbuffer005 As Variant
Dim Varkbuffer008 As Variant
Dim lngHowMany As Long
Dim lngHowManyECis As Long
Dim lngHowManyEDep As Long
Dim lngHowMany005 As Long
Dim lngHowMany008 As Long
Dim SQL_Captura_Ev4 As String
Dim SQL_Marca_Eventos_Env As String
Dim SQL_Estado_Cisterna As String
Dim SQL_Marca_Marca_Comp As String
Dim SQL_Estado_Comp As String
Dim SQL_Ges_B005 As String
Dim SQL_Ges_B008 As String
Dim SQL_Actualiza_hora As String
Dim DA As CDataAccess
Dim RegistroLeid, FechaEvento, NumPed, Particion, NumViaje, MatCisterna, IdPrecinto, Deposito, FechaHasta, FechaDesde As String
Dim Comp1, Comp2, Comp3, Comp4, Comp5, Comp6, Comp7 As String
Dim FeComp1, FeComp2, FeComp3, FeComp4, FeComp5, FeComp6, FeComp7 As String
Dim TiempoMargen, MargenEspera As Long
Dim EventoPr As Integer

SQL_Ges_B005 = "SELECT GESDYR00, GESHRR00 FROM VIAGGES0 WHERE (GESCODEV = '005') AND (GESNUMPD = ?) AND (GESNUMPR = ?)"
SQL_Ges_B008 = "SELECT GESCOMP1, GESCOMP2, GESCOMP3, GESCOMP4, GESCOMP5, GESCOMP6, GESCOMP7 FROM VIAGGES0 WHERE (IDVVT = ?) AND (GESPARAR = 0) AND GESCODEV = '008'"
Set DA = DataConnect.DA
Dim i As Integer


On Error GoTo ErrHandler

    LogMessage "Inicio poceso de Rellenado Compartimentos " & vbCrLf
    XA.value(8, 1) = "En Ejecucion."
    gridProcedures.RefetchCol 1
    ' JG Capturamos todo los eventos  parados
    SQL_Captura_Ev4 = "SELECT        TOP (400) VIAGGES0.GESIDENT, CONVERT(datetime, VIAGGES0.GESDYEV0 + ' ' + VIAGGES0.GESHREV0, 112) AS FechaEv, VIAGGES0.GESNUMPD, VIAGGES0.GESNUMPR, VIAGGES0.IDVVT, " & _
                    "                         ANAGAUT0.AUTIDAM2, ANAGAEC0.AECIDAAM, VIAGGES0.GESCODEV " & _
                    "FROM            ANAGAUT0 LEFT OUTER JOIN " & _
                    "                         ANAGAEC0 ON ANAGAUT0.AUTIDAM2 = ANAGAEC0.AECIDAAM RIGHT OUTER JOIN " & _
                    "                         VIAGVVT0 ON ANAGAUT0.AUTIDAUT = VIAGVVT0.VVTIDAUT RIGHT OUTER JOIN " & _
                    "                         VIAGGES0 ON VIAGVVT0.VVTIDVVT = VIAGGES0.IDVVT " & _
                    "WHERE        (CONVERT(datetime, VIAGGES0.GESDYEV0 + ' ' + VIAGGES0.GESHREV0, 112) < GETDATE()) AND (VIAGGES0.GESTRCEP = 0) " & _
                        "       ORDER BY CONVERT(datetime, dbo.VIAGGES0.GESDYEV0 + ' ' + dbo.VIAGGES0.GESHREV0, 112)"
    If Not DA.QueryCommand(SQL_Captura_Ev4, Varkbuffer, lngHowMany) Then
    End If

For i = 0 To lngHowMany - 1

        RegistroLeid = (Varkbuffer(0, i))
        FechaEvento = (Varkbuffer(1, i))
        NumPed = (Varkbuffer(2, i))
        Particion = Varkbuffer(3, i)
        NumViaje = (Varkbuffer(4, i))
        MatCisterna = (Varkbuffer(5, i))
        IdPrecinto = (Varkbuffer(6, i))
        EventoPr = CInt(Varkbuffer(7, i))
        Comp1 = "0"
        Comp2 = "0"
        Comp3 = "0"
        Comp4 = "0"
        Comp5 = "0"
        Comp6 = "0"
        Comp7 = "0"
        
        ' JG Si no tiene precinto se marca para enviar
        ' Primero miramos que no exista ya el evento 005 en cuyo caso rellenaremos la hora de entrega
        If IsNull(Varkbuffer(6, i)) Or IsEmpty(Varkbuffer(6, i)) Then
            
        ' Si el evento no es 005 miramos si ya ha llegado un 005
            If EventoPr <> 5 Then
            
                    If Not DA.QueryCommand(SQL_Ges_B005, Varkbuffer005, lngHowMany005, NumPed, Particion) Then
                    End If
                
                    If lngHowMany005 > 0 Then
                        SQL_Actualiza_hora = "UPDATE VIAGGES0 SET GESDYR00 ='" & Varkbuffer005(0, 0) & "', GESHRR00 = '" & Varkbuffer005(1, 0) & "' WHERE  (GESIDENT = '" & RegistroLeid & "')"
                        If Not (DA.ExecuteQuery(SQL_Actualiza_hora)) Then
                        End If
                    End If
            End If
        If FechaEvento < Now() Then
            SQL_Marca_Eventos_Env = "UPDATE VIAGGES0 SET GESTRCEP = 1 WHERE(GESIDENT =" & RegistroLeid & ")"
            If Not (DA.ExecuteQuery(SQL_Marca_Eventos_Env)) Then
            End If
        End If
        Else
        
        ' JG Le resto a la hora actual 0.014 que son unos 20 minutos y solo proceso los que sean menores de esa hora
        ' Establecemos el margen de tiempo parado evento en dias
        ' Establecemos en tiempomargen en minutos el margen de error del evento con el estado de la cisterna
            MargenEspera = 15 / 60 / 60 / 24
            If (FechaEvento < (Now() - MargenEspera)) Or (EventoPr <> 4) Then
            TiempoMargen = 240
            ' Buscamos en la tabla histórica cual es el último estado de la cisterna
            SQL_Estado_Cisterna = "SELECT     AECIDAAM, AECEPDP1, AECFEES1, AECEPDP2, AECFEES2, AECEPDP3, AECFEES3, AECEPDP4, AECFEES4, AECEPDP5, AECFEES5, AECEPDP6, " & _
                                    "                      AECFEES6, AECEPDP7, AECFEES7, AECEPDP8, AECFEES8  FROM ANAGAEC0 WHERE     (AECIDAAM = ?)"

            If Not DA.QueryCommand(SQL_Estado_Cisterna, VarkbufferECis, lngHowManyECis, MatCisterna) Then
            End If
                If lngHowManyECis > 0 Then
                ' Miramos hasta 5 minutos despues por si el conductor ha mandado la salida antes de cargar
                FechaHasta = FechaEvento + (2 / 60 / 24)
                FechaDesde = (FechaEvento - (TiempoMargen / 60 / 24))
                
            ' Miramos compartimento 1

                        If Not IsNull(VarkbufferECis(1, 0)) Then

                            Deposito = "01"
                            FeComp1 = VarkbufferECis(2, 0)
                            ' Buscamos cual es el estado último del compartimento en el margen establecido
                            SQL_Estado_Comp = "SELECT     TOP 100 PERCENT dbo.ANAGAEP0.AEPESCEP FROM OESIA.dbo.EventosCisternasMBFlotas LEFT OUTER JOIN " & _
                            "dbo.ANAGAEP0 ON OESIA.dbo.EventosCisternasMBFlotas.Subestructura2 = dbo.ANAGAEP0.AEPIDEPD COLLATE Modern_Spanish_CI_AS " & _
                            "WHERE     (OESIA.dbo.EventosCisternasMBFlotas.FechaEvento < CONVERT(DATETIME, '" & FechaHasta & "', 103) AND " & _
                            "OESIA.dbo.EventosCisternasMBFlotas.FechaEvento > CONVERT(DATETIME, '" & FechaDesde & "', 103)) AND " & _
                            "(OESIA.dbo.EventosCisternasMBFlotas.CodigoVehiculo = ?) AND (OESIA.dbo.EventosCisternasMBFlotas.Deposito = ?) AND " & _
                            "(dbo.ANAGAEP0.AEPESCEP IS NOT NULL) ORDER BY OESIA.dbo.EventosCisternasMBFlotas.IdRegistro DESC "
    
                            If Not DA.QueryCommand(SQL_Estado_Comp, VarkbufferEDep, lngHowManyEDep, IdPrecinto, Deposito) Then
                            End If
    
                            If lngHowManyEDep > 0 Then
                            Comp1 = VarkbufferEDep(0, 0)
                            End If
                            ' Si el compartimento tiene estado de desprecintado en el momento de la lectura buscamos el ultimo en la bd y si tambien es desprecintado no lo mandamos
                            If Comp1 = "0" Then
                                    Comp1 = VarkbufferECis(1, 0)
                                    If (Round((FechaHasta - FeComp1) * 24 * 60, 0)) > 0 And (Round((FechaHasta - FeComp1) * 24 * 60, 0)) < 86400 Then
                                    Else
                                    Comp1 = "0"
                                    End If
                            End If
                            If Comp1 > "22" Then Comp1 = "0"
                        End If

                ' Miramos compartimento 2

                    If Not IsNull(VarkbufferECis(3, 0)) Then
                        Deposito = "02"
                        FeComp2 = VarkbufferECis(4, 0)

                        SQL_Estado_Comp = "SELECT     TOP 100 PERCENT dbo.ANAGAEP0.AEPESCEP FROM OESIA.dbo.EventosCisternasMBFlotas LEFT OUTER JOIN " & _
                        "dbo.ANAGAEP0 ON OESIA.dbo.EventosCisternasMBFlotas.Subestructura2 = dbo.ANAGAEP0.AEPIDEPD COLLATE Modern_Spanish_CI_AS " & _
                        "WHERE     (OESIA.dbo.EventosCisternasMBFlotas.FechaEvento < CONVERT(DATETIME, '" & FechaHasta & "', 103) AND " & _
                        "OESIA.dbo.EventosCisternasMBFlotas.FechaEvento > CONVERT(DATETIME, '" & FechaDesde & "', 103)) AND " & _
                        "(OESIA.dbo.EventosCisternasMBFlotas.CodigoVehiculo = ?) AND (OESIA.dbo.EventosCisternasMBFlotas.Deposito = ?) AND " & _
                        "(dbo.ANAGAEP0.AEPESCEP IS NOT NULL) ORDER BY OESIA.dbo.EventosCisternasMBFlotas.IdRegistro DESC "

                        If Not DA.QueryCommand(SQL_Estado_Comp, VarkbufferEDep, lngHowManyEDep, IdPrecinto, Deposito) Then
                        End If

                        If lngHowManyEDep > 0 Then
                        Comp2 = VarkbufferEDep(0, 0)
                        End If

                            ' Si el compartimento tiene estado de desprecintado en el momento de la lectura buscamos el ultimo en la bd y si tambien es desprecintado no lo mandamos
                            If Comp2 = "0" Then
                                    Comp2 = VarkbufferECis(3, 0)
                                    If (Round((FechaHasta - FeComp2) * 24 * 60, 0)) > 0 And (Round((FechaHasta - FeComp2) * 24 * 60, 0)) < 86400 Then
                                    Else
                                    Comp2 = "0"
                                    End If
                            End If
                            If Comp2 > "22" Then Comp2 = "0"
                    End If

                ' Miramos compartimento 3

                    If Not IsNull(VarkbufferECis(5, 0)) Then
                        Deposito = "03"
                        FeComp3 = VarkbufferECis(6, 0)
                        SQL_Estado_Comp = "SELECT     TOP 100 PERCENT dbo.ANAGAEP0.AEPESCEP FROM OESIA.dbo.EventosCisternasMBFlotas LEFT OUTER JOIN " & _
                        "dbo.ANAGAEP0 ON OESIA.dbo.EventosCisternasMBFlotas.Subestructura2 = dbo.ANAGAEP0.AEPIDEPD COLLATE Modern_Spanish_CI_AS " & _
                        "WHERE     (OESIA.dbo.EventosCisternasMBFlotas.FechaEvento < CONVERT(DATETIME, '" & FechaHasta & "', 103) AND " & _
                        "OESIA.dbo.EventosCisternasMBFlotas.FechaEvento > CONVERT(DATETIME, '" & FechaDesde & "', 103)) AND " & _
                        "(OESIA.dbo.EventosCisternasMBFlotas.CodigoVehiculo = ?) AND (OESIA.dbo.EventosCisternasMBFlotas.Deposito = ?) AND " & _
                        "(dbo.ANAGAEP0.AEPESCEP IS NOT NULL) ORDER BY OESIA.dbo.EventosCisternasMBFlotas.IdRegistro DESC "

                        If Not DA.QueryCommand(SQL_Estado_Comp, VarkbufferEDep, lngHowManyEDep, IdPrecinto, Deposito) Then
                        End If

                        If lngHowManyEDep > 0 Then
                        Comp3 = VarkbufferEDep(0, 0)
                        End If

                            ' Si el compartimento tiene estado de desprecintado en el momento de la lectura buscamos el ultimo en la bd y si tambien es desprecintado no lo mandamos
                            If Comp3 = "0" Then
                                    Comp3 = VarkbufferECis(5, 0)
                                    If (Round((FechaHasta - FeComp3) * 24 * 60, 0)) > 0 And (Round((FechaHasta - FeComp3) * 24 * 60, 0)) < 86400 Then
                                    Else
                                    Comp3 = "0"
                                    End If
                            End If
                            If Comp3 > "22" Then Comp3 = "0"

                    End If

                ' Miramos compartimento 4


                    If Not IsNull(VarkbufferECis(7, 0)) Then
                        Deposito = "04"
                        FeComp4 = VarkbufferECis(8, 0)

                        SQL_Estado_Comp = "SELECT     TOP 100 PERCENT dbo.ANAGAEP0.AEPESCEP FROM OESIA.dbo.EventosCisternasMBFlotas LEFT OUTER JOIN " & _
                        "dbo.ANAGAEP0 ON OESIA.dbo.EventosCisternasMBFlotas.Subestructura2 = dbo.ANAGAEP0.AEPIDEPD COLLATE Modern_Spanish_CI_AS " & _
                        "WHERE     (OESIA.dbo.EventosCisternasMBFlotas.FechaEvento < CONVERT(DATETIME, '" & FechaHasta & "', 103) AND " & _
                        "OESIA.dbo.EventosCisternasMBFlotas.FechaEvento > CONVERT(DATETIME, '" & FechaDesde & "', 103)) AND " & _
                        "(OESIA.dbo.EventosCisternasMBFlotas.CodigoVehiculo = ?) AND (OESIA.dbo.EventosCisternasMBFlotas.Deposito = ?) AND " & _
                        "(dbo.ANAGAEP0.AEPESCEP IS NOT NULL) ORDER BY OESIA.dbo.EventosCisternasMBFlotas.IdRegistro DESC "

                        If Not DA.QueryCommand(SQL_Estado_Comp, VarkbufferEDep, lngHowManyEDep, IdPrecinto, Deposito) Then
                        End If

                        If lngHowManyEDep > 0 Then
                        Comp4 = VarkbufferEDep(0, 0)
                        End If

                            ' Si el compartimento tiene estado de desprecintado en el momento de la lectura buscamos el ultimo en la bd y si tambien es desprecintado no lo mandamos
                            If Comp4 = "0" Then
    
                                    Comp4 = VarkbufferECis(7, 0)
                                    If (Round((FechaHasta - FeComp4) * 24 * 60, 0)) > 0 And (Round((FechaHasta - FeComp4) * 24 * 60, 0)) < 86400 Then
                                    Else
                                    Comp4 = "0"
                                    End If
                            End If
                            If Comp4 > "22" Then Comp4 = "0"
                        End If


                ' Miramos compartimento 5

                       If Not IsNull(VarkbufferECis(9, 0)) Then
                        Deposito = "05"
                        FeComp5 = VarkbufferECis(10, 0)

                        SQL_Estado_Comp = "SELECT     TOP 100 PERCENT dbo.ANAGAEP0.AEPESCEP FROM OESIA.dbo.EventosCisternasMBFlotas LEFT OUTER JOIN " & _
                        "dbo.ANAGAEP0 ON OESIA.dbo.EventosCisternasMBFlotas.Subestructura2 = dbo.ANAGAEP0.AEPIDEPD COLLATE Modern_Spanish_CI_AS " & _
                        "WHERE     (OESIA.dbo.EventosCisternasMBFlotas.FechaEvento < CONVERT(DATETIME, '" & FechaHasta & "', 103) AND " & _
                        "OESIA.dbo.EventosCisternasMBFlotas.FechaEvento > CONVERT(DATETIME, '" & FechaDesde & "', 103)) AND " & _
                        "(OESIA.dbo.EventosCisternasMBFlotas.CodigoVehiculo = ?) AND (OESIA.dbo.EventosCisternasMBFlotas.Deposito = ?) AND " & _
                        "(dbo.ANAGAEP0.AEPESCEP IS NOT NULL) ORDER BY OESIA.dbo.EventosCisternasMBFlotas.IdRegistro DESC "

                        If Not DA.QueryCommand(SQL_Estado_Comp, VarkbufferEDep, lngHowManyEDep, IdPrecinto, Deposito) Then
                        End If

                        If lngHowManyEDep > 0 Then
                        Comp5 = VarkbufferEDep(0, 0)
                        End If

                            ' Si el compartimento tiene estado de desprecintado en el momento de la lectura buscamos el ultimo en la bd y si tambien es desprecintado no lo mandamos
                            If Comp5 = "0" Then
    
                                    Comp5 = VarkbufferECis(9, 0)
                                    If (Round((FechaHasta - FeComp5) * 24 * 60, 0)) > 0 And (Round((FechaHasta - FeComp5) * 24 * 60, 0)) < 86400 Then
                                    Else
                                    Comp5 = "0"
                                    End If
                            End If
                            If Comp5 > "22" Then Comp5 = "0"
                        End If


                ' Miramos compartimento 6

                    If Not IsNull(VarkbufferECis(11, 0)) Then
                        Deposito = "06"
                        FeComp6 = VarkbufferECis(12, 0)

                        SQL_Estado_Comp = "SELECT     TOP 100 PERCENT dbo.ANAGAEP0.AEPESCEP FROM OESIA.dbo.EventosCisternasMBFlotas LEFT OUTER JOIN " & _
                        "dbo.ANAGAEP0 ON OESIA.dbo.EventosCisternasMBFlotas.Subestructura2 = dbo.ANAGAEP0.AEPIDEPD COLLATE Modern_Spanish_CI_AS " & _
                        "WHERE     (OESIA.dbo.EventosCisternasMBFlotas.FechaEvento < CONVERT(DATETIME, '" & FechaHasta & "', 103) AND " & _
                        "OESIA.dbo.EventosCisternasMBFlotas.FechaEvento > CONVERT(DATETIME, '" & FechaDesde & "', 103)) AND " & _
                        "(OESIA.dbo.EventosCisternasMBFlotas.CodigoVehiculo = ?) AND (OESIA.dbo.EventosCisternasMBFlotas.Deposito = ?) AND " & _
                        "(dbo.ANAGAEP0.AEPESCEP IS NOT NULL) ORDER BY OESIA.dbo.EventosCisternasMBFlotas.IdRegistro DESC "

                        If Not DA.QueryCommand(SQL_Estado_Comp, VarkbufferEDep, lngHowManyEDep, IdPrecinto, Deposito) Then
                        End If

                        If lngHowManyEDep > 0 Then
                        Comp6 = VarkbufferEDep(0, 0)
                        End If

                            ' Si el compartimento tiene estado de desprecintado en el momento de la lectura buscamos el ultimo en la bd y si tambien es desprecintado no lo mandamos
                            If Comp6 = "0" Then
                                    
                                    Comp6 = VarkbufferECis(11, 0)
                                    If (Round((FechaHasta - FeComp6) * 24 * 60, 0)) > 0 And (Round((FechaHasta - FeComp6) * 24 * 60, 0)) < 86400 Then
                                    Else
                                    Comp6 = "0"
                                    End If
                            End If
                            If Comp6 > "22" Then Comp6 = "0"
                        End If



                ' Miramos compartimento 7

                    If Not IsNull(VarkbufferECis(13, 0)) Then
                        Deposito = "07"
                        FeComp7 = VarkbufferECis(12, 0)

                        SQL_Estado_Comp = "SELECT     TOP 100 PERCENT dbo.ANAGAEP0.AEPESCEP FROM OESIA.dbo.EventosCisternasMBFlotas LEFT OUTER JOIN " & _
                        "dbo.ANAGAEP0 ON OESIA.dbo.EventosCisternasMBFlotas.Subestructura2 = dbo.ANAGAEP0.AEPIDEPD COLLATE Modern_Spanish_CI_AS " & _
                        "WHERE     (OESIA.dbo.EventosCisternasMBFlotas.FechaEvento < CONVERT(DATETIME, '" & FechaHasta & "', 103) AND " & _
                        "OESIA.dbo.EventosCisternasMBFlotas.FechaEvento > CONVERT(DATETIME, '" & FechaDesde & "', 103)) AND " & _
                        "(OESIA.dbo.EventosCisternasMBFlotas.CodigoVehiculo = ?) AND (OESIA.dbo.EventosCisternasMBFlotas.Deposito = ?) AND " & _
                        "(dbo.ANAGAEP0.AEPESCEP IS NOT NULL) ORDER BY OESIA.dbo.EventosCisternasMBFlotas.IdRegistro DESC "

                        If Not DA.QueryCommand(SQL_Estado_Comp, VarkbufferEDep, lngHowManyEDep, IdPrecinto, Deposito) Then
                        End If

                        If lngHowManyEDep > 0 Then
                        Comp7 = VarkbufferEDep(0, 0)
                        End If

                            ' Si el compartimento tiene estado de desprecintado en el momento de la lectura buscamos el ultimo en la bd y si tambien es desprecintado no lo mandamos
                            If Comp7 = "0" Then
                                    
                                    Comp7 = VarkbufferECis(13, 0)
                                    If (Round((FechaHasta - FeComp7) * 24 * 60, 0)) > 0 And (Round((FechaHasta - FeComp7) * 24 * 60, 0)) < 86400 Then
                                    Else
                                    Comp7 = "0"
                                    End If
                            End If
                            If Comp7 > "22" Then Comp7 = "0"
                        End If


                End If
                ' Si no es un evento 005 miramos si ya se ha llegado
                If EventoPr <> 5 Then
                    If Not DA.QueryCommand(SQL_Ges_B005, Varkbuffer005, lngHowMany005, NumPed, Particion) Then
                    End If
                
                    If lngHowMany005 > 0 Then
                        SQL_Actualiza_hora = "UPDATE VIAGGES0 SET GESDYR00 ='" & Varkbuffer005(0, 0) & "', GESHRR00 = '" & Varkbuffer005(1, 0) & "' WHERE  (GESIDENT = '" & RegistroLeid & "')"
                        If Not (DA.ExecuteQuery(SQL_Actualiza_hora)) Then
                        End If
                    End If
                End If
                
                ' Miramos si hay un evento 0008 enviado en cuyo caso cambiamos los estado de los tanques a los que figuran en el 008
                    If Not DA.QueryCommand(SQL_Ges_B008, Varkbuffer008, lngHowMany008, NumViaje) Then
                    End If
                
                    If lngHowMany008 > 0 Then
                       If Not IsNull(Varkbuffer008(0, 0)) Then Comp1 = Varkbuffer008(0, 0)
                       If Not IsNull(Varkbuffer008(1, 0)) Then Comp2 = Varkbuffer008(1, 0)
                       If Not IsNull(Varkbuffer008(2, 0)) Then Comp3 = Varkbuffer008(2, 0)
                       If Not IsNull(Varkbuffer008(3, 0)) Then Comp4 = Varkbuffer008(3, 0)
                       If Not IsNull(Varkbuffer008(4, 0)) Then Comp5 = Varkbuffer008(4, 0)
                       If Not IsNull(Varkbuffer008(5, 0)) Then Comp6 = Varkbuffer008(5, 0)
                       If Not IsNull(Varkbuffer008(6, 0)) Then Comp7 = Varkbuffer008(6, 0)
                    End If

                
                SQL_Marca_Marca_Comp = "UPDATE VIAGGES0 SET GESCOMP1 = '" & Comp1 & "',GESCOMP2 = '" & Comp2 & "',GESCOMP3 = '" & Comp3 & "',GESCOMP4 = '" & Comp4 & "',GESCOMP5 = '" & Comp5 & "',GESCOMP6 = '" & Comp6 & "',GESCOMP7 = '" & Comp7 & "' , GESTRCEP = 1 WHERE(GESIDENT =" & RegistroLeid & ")"
                If Not (DA.ExecuteQuery(SQL_Marca_Marca_Comp)) Then
                
                End If

                End If
        End If

Next i

GoTo EndProc

ErrHandler:
    
    XA.value(8, 1) = "En Error."
    gridProcedures.RefetchCol 1
    Exit Function
    
EndProc:

    LogMessage "Fin poceso de Rellenado Compartimentos " & vbCrLf
    XA.value(8, 1) = "En Espera"
    gridProcedures.RefetchCol 1


End Function
Public Function GeneraGes_EUT()
Dim DA As CDataAccess

    LogMessage "Inicio poceso de Generación GES Eutrasur " & vbCrLf

   Set DA = DataConnect.DA

    Dim i As Integer
    Dim Varkbuffer As Variant
    Dim lngHowMany As Long
    Dim SQL_GES_recuperar As String
   
    SQL_GES_recuperar = "SELECT  * From VIAGGES0 WHERE (GESSEND0 = 0) AND GESCODTR='004' AND GESTRCEP = '1' AND GESPARAR = 0 AND (GESNUMPD IS NOT NULL)"
       
'   Dim ERNstrTempFileName As String
'   ERNstrTempFileName = "GES_"

    Dim NLote As String
    Dim FechaGL As String
    Dim HoraGL As String
    Dim NCampos As Integer
    
    XA.value(2, 1) = "En Ejecucion"
    gridProcedures.RefetchCol 1

    
     If Not DA.QueryCommand(SQL_GES_recuperar, Varkbuffer, lngHowMany) Then
        'objLogger.logError "Query SQL_GES_recuperar."
        'GoTo ErrHandler
   End If
   
   If (lngHowMany > 0) Then
    
        ' SELECIONAR EL NUMERO DE LOTE
               
             With DataClass.DS.DA
                    NLote = .GetCounter("anaganu0", "anuprogr", "anuidanu", "107", 0)
                    If NLote = -1 Then
        ''                ErrCode = InvalidCode
        ''                GoTo ErrHandler
                    End If
                    If NLote > 9999999999# Then
                        NLote = 1
                    End If
            End With
'           NLote = Format$(Time, "hhmmss")
                
            FechaGL = Format$(Date, "YYYYMMDD")
            HoraGL = Format$(Hour(Now), "00") & ":" & Format$(Minute(Now), "00") & ":" & Format$(Second(Now), "00")
        
            Dim ficheroeut As String
           
           If Len(mskFileNameEUT.Text) < 2 Then
           ficheroeut = MaskEdUnidad.Text & ":\cepsa\eutrasur\salida"
           Else
           ficheroeut = mskFileNameEUT.Text
           End If
           
          If Not mobjOutFile.OpenFile(ficheroeut + "\GES_N" + NombreGes(NLote) + ".DAT", 1) Then
'            Dim nombrearch  As String
'            nombrearch = Format$(Date, "ddmmyy") & Format$(Time, "hhmmss")
'
'            If Not mobjOutFile.OpenFile(ficheroeut + "\GES_EUT" + nombrearch + ".DAT", 1) Then
'
'                'ErrCode = FileNotFound
'                'GoTo ErrHandler
           End If
           
           
          
            
            NCampos = lngHowMany
                   
        'comprobamos que tenemos algo que enviar en el GES
        If lngHowMany > 0 Then
                
          For i = 0 To UBound(Varkbuffer, 2)
               With mobjOutFile
               
               .ClearFields
               
                    .Item("1").value = "3369" ' "0" + Varkbuffer(1, i) POner Oilogist
                    'generados cuando se genera el GES
                    .Item("2").value = NLote 'numero de LOTE
                    .Item("3").value = FechaGL
                    .Item("4").value = HoraGL
                    .Item("5").value = NCampos
                    '----------
                    .Item("6").value = Varkbuffer(6, i)
                 'Numero de partición "00" si no hay.
                     If (Len(Varkbuffer(7, i)) < 2) Then
                    .Item("7").value = "00"
                     Else
                    .Item("7").value = Varkbuffer(7, i)
                     End If
                    .Item("8").value = Varkbuffer(8, i)
                    'Tipo de coordenada en nuestro caso siempre 01
                    .Item("9").value = "01"
                    .Item("10").value = Varkbuffer(10, i)
                    .Item("11").value = Varkbuffer(11, i)
                    .Item("12").value = Varkbuffer(12, i)
                    .Item("13").value = Varkbuffer(13, i)
                    .Item("14").value = Varkbuffer(14, i)
                    'Unidad de medida, ("L" o "K")
                    If (Varkbuffer(15, i) = "K") Then
                    .Item("15").value = Varkbuffer(15, i)
                    Else
                    .Item("15").value = "L"
                    End If
                    
                    .Item("16").value = Varkbuffer(16, i)
                   
                     'hora actualizada
                    If (Varkbuffer(17, i) = "00:00") Then
                     .Item("17").value = Varkbuffer(23, i)
                    Else
                     .Item("17").value = Varkbuffer(17, i)
                    End If
                    
                    If (Varkbuffer(18, i) = "19000101") Then
                    'fecha actualizada.
                     .Item("18").value = Varkbuffer(24, i)
                    Else
                     .Item("18").value = Varkbuffer(18, i)
                    End If
                   
                    .Item("19").value = Varkbuffer(19, i)
                    .Item("20").value = Varkbuffer(20, i)
'                    .Item("21").value = Varkbuffer(21, i)
'                    .Item("22").value = Varkbuffer(22, i)
'                    .Item("23").value = Varkbuffer(23, i)
'                    .Item("24").value = Varkbuffer(24, i)
                    .Item("21").value = Varkbuffer(25, i)
                    .Item("22").value = Varkbuffer(26, i)
                    .Item("23").value = Varkbuffer(27, i)
                    .Item("24").value = Varkbuffer(28, i)
                    .Item("25").value = Varkbuffer(29, i)
                      If (IsNull(Varkbuffer(30, i))) Then
                    Else
                    .Item("26").value = Replace(Varkbuffer(30, i), ",", ".")
                    End If
                    
                    If (IsNull(Varkbuffer(31, i))) Then
                    Else
                    .Item("27").value = Replace(Varkbuffer(31, i), ",", ".")
                    End If
                     
                    If (IsNull(Varkbuffer(32, i))) Then
                    Else
                    .Item("28").value = Replace(Varkbuffer(32, i), ",", ".")
                    End If
                    
                    .Item("29").value = Varkbuffer(33, i)
                    .Item("30").value = Varkbuffer(34, i)
                    'si el evento es 999 (actualización de horario) no rellenar el codigo de evento.
                    If (Trim(Varkbuffer(35, i)) = "999") Then
                     .Item("31").value = ""
                    Else
                     .Item("31").value = Varkbuffer(35, i)
                    End If
                    .Item("32").value = Varkbuffer(36, i)
                    .Item("33").value = Varkbuffer(37, i)
                    .Item("34").value = Varkbuffer(38, i)
                    .Item("35").value = Varkbuffer(39, i)
                    .Item("36").value = Varkbuffer(40, i)
                    .Item("37").value = Varkbuffer(41, i)
                    .Item("38").value = Varkbuffer(42, i)
                    .Item("39").value = Varkbuffer(43, i)
                    .Item("40").value = Varkbuffer(44, i)
                    .Item("41").value = Varkbuffer(45, i)
                    .Item("42").value = Varkbuffer(46, i)
                    .Item("43").value = 0 'varKBuffer(47, i) PACTAR CLIENTE = 0
                    .Item("44").value = Varkbuffer(22, i) ' LA MISMA HORA QUE PARA EL TRANSPORTISTA' HORA PREVISTA DE CARGA EN FACTORIA
                    
                  'NUEVOS CAMPOS
                    .Item("45").value = Varkbuffer(21, i) 'fecha prevista carga
                    Dim rowbufmat As Variant
                    Dim longmat As Long
                    
                    
                Dim sql_r As String
                          sql_r = "SELECT     ANAGAAM0_1.AAMMTCEP AS TRact, ANAGAAM0.AAMMTCEP AS CIst " & _
                         " FROM         ANAGAAM0 RIGHT OUTER JOIN " & _
                         " ANAGAUT0 ON ANAGAAM0.AAMIDAAM = ANAGAUT0.AUTIDAM2 LEFT OUTER JOIN " & _
                         " ANAGAAM0 AS ANAGAAM0_1 ON ANAGAUT0.AUTIDAM1 = ANAGAAM0_1.AAMIDAAM RIGHT OUTER JOIN " & _
                         " VIAGVVT0 ON ANAGAUT0.AUTIDAUT = VIAGVVT0.VVTIDAUT " & _
                         "WHERE     (VIAGVVT0.VVTIDVVT = ?) "




                    If Not DA.QueryCommand(sql_r, rowbufmat, longmat, CStr(Varkbuffer(54, i))) Then
                        'ERROR
                         GoTo ErrHandler
                    End If
                    If longmat > 0 Then
                    .Item("46").value = rowbufmat(0, 0) 'mat trac
                    .Item("47").value = rowbufmat(1, 0) 'mat cis
                    End If
                    .Item("48").value = "" 'mat 3
                    .Item("49").value = Varkbuffer(56, i) ' Motivo INC
                    .Item("50").value = Varkbuffer(57, i) ' Cantidadsss
                    
                    Dim PROD As Variant
                    Dim longprod As Long
                    


                If Not IsNull(Varkbuffer(58, i)) And Varkbuffer(58, i) <> "" Then
                    If (Trim(Varkbuffer(35, i)) = "011") Then
                    .Item("51").value = Varkbuffer(58, i)
                    Else
                    .Item("51").value = ""
                    End If
                End If
                    
                If Not IsNull(Varkbuffer(59, i)) And Varkbuffer(59, i) <> "" Then
                    If (Trim(Varkbuffer(35, i)) = "011") Then
                    .Item("52").value = Varkbuffer(59, i)
                    Else
                    .Item("52").value = ""
                    End If
                End If
                
                
                If Varkbuffer(60, i) = True Then
                    .Item("53").value = "S" 'resuelto medios propios?
                Else
                       If (Trim(Varkbuffer(35, i)) = "012") Then
                             .Item("53").value = "N"
                        Else
                            .Item("53").value = ""
                        End If
                End If
                
                 ' Tiempo paraliza
                If Not IsNull(Varkbuffer(61, i)) Then
                        If (Trim(Varkbuffer(35, i)) = "019") Or (Trim(Varkbuffer(35, i)) = "020") Then
                             .Item("54").value = Replace(Format$(CStr(Varkbuffer(61, i)), "000.0"), ",", ".")
                        Else
                            .Item("54").value = ""
                        End If
                End If
                
                ' Destinatario desvio
                If Not IsNull(Varkbuffer(62, i)) Then
                        If (Trim(Varkbuffer(35, i)) = "017") Then
                             .Item("55").value = Replace(Format$(CStr(Varkbuffer(62, i)), "000.0"), ",", ".")
                        Else
                            .Item("55").value = ""
                        End If
                End If
                
                ' JG 25/11/21 Nuevo campo observaciones
                
                 If Not IsNull(Varkbuffer(65, i)) Then
                            .Item("56").value = CStr(Varkbuffer(65, i))
                 Else
                            .Item("56").value = ""
                 End If
                
               
     
        
                End With
              mobjOutFile.WriteLine
            Next i
        
           mobjOutFile.CloseFile
            
            'Updatear los capos generados en el GES
            Dim SQL_UPDATE_GESENVIADOS As String
            Dim SQL_UPDATE_GES As String
            SQL_UPDATE_GESENVIADOS = "UPDATE VIAGGES0 SET GESSEND0 = 1, GESNUMLO =? ,GESDAYDY = ?, GESDATHR = ? ,GESNUMRE = ?  WHERE GESIDENT = ?" 'LOTE,FECHA,HORA,NUMREGISTROS WHERE ID
        
            Dim IDGES As Long
            
            For i = 0 To UBound(Varkbuffer, 2)
                    IDGES = Varkbuffer(0, i)
                    If Not DA.ExecuteCommand(SQL_UPDATE_GESENVIADOS, NLote, FechaGL, HoraGL, NCampos, IDGES) Then
                    End If
            Next i
        End If
  End If
ErrHandler:

    XA.value(2, 1) = "Error"
    gridProcedures.RefetchCol 1

EndProc:

    XA.value(2, 1) = "En Espera"
    gridProcedures.RefetchCol 1
    
   
End Function
Private Sub Form_Unload(Cancel As Integer)
    gmRegistry.SaveSetting "Options", REGISTRY_KEY_NAME_PATH_TPH, Me.mskFileNameTPH, True
    gmRegistry.SaveSetting "Options", REGISTRY_KEY_NAME_PATH_EUT, Me.mskFileNameEUT, True
    gmRegistry.SaveSetting "Options", REGISTRY_KEY_NAME_LAST_SMS, mskLastSMS.Text, True
    gmRegistry.SaveSetting "Options", REGISTRY_KEY_NAME_UNIDAD, MaskEdUnidad.Text, True
    gmRegistry.SaveSetting "Options", REGISTRY_KEY_NAME_HORAS, MaskEdHoras.Text, True
   
    Set XA = Nothing

    Set frmGeneraGes = Nothing
End Sub

Private Sub gridProcedures_ButtonClick(ByVal ColIndex As Integer)
    If ColIndex = 2 Then
        Select Case gridProcedures.bookmark
            Case 0
                Execute_ScanEventsPalm_nuevo
            Case 1
                GeneraGes_TPH
            Case 2
                GeneraGes_EUT
            Case 3
                RecepcionEcoges
            Case 4
                Captura_SMS
            Case 5
                Captura_albaranes
            Case 6
                Captura_prec_oesia
            Case 7
                Captura_prec_thaumat
            Case 8
                Rellena_Compartimentos
            Case 9
                Captura_desprecintados
                
        End Select
    End If
End Sub
Private Sub LogMessage(message As String)
    textLog.Text = textLog.Text & Format(Now(), "Hh:Nn.Ss") & " - " & message
    textLog.SelStart = Len(textLog.Text)
End Sub
Private Function Execute_ScanEventsPalm_nuevo() As Boolean
On Error GoTo ErrHandler
    Dim varLastID As Variant
    Dim varMinuti As Variant
    Dim DA As CDataAccess
    Set DA = DataConnect.DA

    Dim KVE As CSatellitareEvento
    
 '--SM -JG Detección de ejecución de proccesos
        Dim ret As Boolean
        Dim x As CTextFile
        Set x = New CTextFile
        ret = x.OpenFile(MaskEdUnidad.Text & ":\Resoil2\Alarmas\PALM.txt", 1)
        ret = x.CloseFile()

    Set KVE = New CSatellitareEvento
    With KVE
        Set .DS.DA = DataConnect.DA
        Set .DS.ErrH = DataConnect.DA.ErrH

'        KVE.Logger.LogFileName = gmstrServerLogFileName
'        KVE.Logger.enableLevel (strLogLevels)

        XA.value(0, 1) = "En Ejecucion"
        gridProcedures.RefetchCol 1

                If mblnRunOk Then
                    mblnRunOk = False
                    LogMessage "Inicio poceso de escaneo eventos de Palms: " & vbCrLf

                    If Not KVE.ScanEvents_Palm_nuevo() Then
                        GoTo ErrHandler
                    End If

                    mblnRunOk = True
                End If

    End With

   

    Execute_ScanEventsPalm_nuevo = True
    Set KVE = Nothing
    GoTo EndProc

ErrHandler:
    Set KVE = Nothing

    mblnRunOk = True
    XA.value(0, 1) = "Error"
    gridProcedures.RefetchCol 1

EndProc:

Dim rowbuff As Variant
Dim lngnocap As Long

'Comprobamos que los eventos se estén pegando correctamente, si no dejamos un fichero en el directorio alarmas.
If DA.QueryCommand("Select * From EventosMBFlotasManual Where Capturado = 0 ", rowbuff, lngnocap) Then
    If (lngnocap > 5) Then
        ret = x.OpenFile(MaskEdUnidad.Text & ":\Resoil2\Alarmas\CAPTURADOS.txt", 1)
        ret = x.CloseFile()
    End If
End If

'Comprobamos que no existen eventos parados en la tabla VIAGES 0
If DA.QueryCommand("SELECT * FROM dbo.VIAGGES0 WHERE (GESPARAR = 1)", rowbuff, lngnocap) Then
    If (lngnocap > 1) Then
        ret = x.OpenFile(MaskEdUnidad.Text & ":\Resoil2\Alarmas\PARADOS.txt", 1)
        ret = x.CloseFile()
    End If
End If

'Comprobamos que no existen viajes programados para una UDT sin GPS
'If DA.QueryCommand("SELECT dbo.ANAGAAM0.AAMIDDEU, dbo.ANAGAAM0.AAMIDAAM  FROM dbo.VIAGVVT0 INNER JOIN dbo.ANAGAAM0 ON dbo.VIAGVVT0.VVTIDAM1 = dbo.ANAGAAM0.AAMIDAAM " & _
'" WHERE     (dbo.VIAGVVT0.VVTDTRIF > GETDATE() - 1) AND (dbo.VIAGVVT0.VVTIDASC = '038') AND (dbo.ANAGAAM0.AAMIDDEU = '' OR dbo.ANAGAAM0.AAMIDDEU IS NULL)", rowbuff, lngnocap) Then
'    If (lngnocap > 0) Then
'        ret = x.OpenFile(MaskEdUnidad.Text & ":\Resoil2\Alarmas\SINGPS.TXT", 1)
'        ret = x.CloseFile()
'    End If
'End If

'Comprobamos que no existen Gps Falsos
'If DA.QueryCommand("SELECT     dbo.ANAGAAM0.AAMIDAAM, dbo.ANAGAAM0.AAMIDDEU FROM dbo.PosicionesOesia RIGHT OUTER JOIN dbo.ANAGAAM0 ON dbo.PosicionesOesia.IDVEHIC = dbo.ANAGAAM0.AAMIDDEU " & _
'"WHERE     (dbo.PosicionesOesia.IDVEHIC IS NULL) AND (dbo.ANAGAAM0.AAMIDDEU <> '') ", rowbuff, lngnocap) Then
'    If (lngnocap > 0) Then
'        ret = x.OpenFile(MaskEdUnidad.Text & ":\Resoil2\Alarmas\GPSFALSO.TXT", 1)
'        ret = x.CloseFile()
'    End If
'End If

XA.value(0, 1) = "En espera"
  gridProcedures.RefetchCol 1

End Function

Private Function ExecuteProcedure()
    
    If Len(textLog.Text) > 10000 Then
        textLog.Text = ""
    End If
    
    
    Execute_ScanEventsPalm_nuevo

    GeneraGes_TPH

    GeneraGes_EUT

    RecepcionEcoges

    Captura_SMS

    Captura_albaranes

    Captura_prec_oesia
    
    Captura_prec_thaumat

    Rellena_Compartimentos
    
    If Chkdesp.value = 1 Then
    Captura_desprecintados Captura_desprecintados
    End If
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




