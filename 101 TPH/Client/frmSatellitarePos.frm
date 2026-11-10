VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmSatellitarePos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Procedure per l'aggiornamento dei dati satellitare"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8745
   ScaleWidth      =   11220
   Begin VB.Frame Frame3 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   8
      Top             =   7800
      Width           =   11235
      Begin VB.CommandButton Command2 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   9480
         Picture         =   "frmSatellitarePos.frx":0000
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
         Picture         =   "frmSatellitarePos.frx":0102
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
         Picture         =   "frmSatellitarePos.frx":0204
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
         Picture         =   "frmSatellitarePos.frx":0306
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
      Height          =   2955
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   4800
      Width           =   11175
   End
   Begin VB.Frame Frame1 
      Caption         =   " Parametri "
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      Begin VB.Timer timerExecuteProcedure 
         Enabled         =   0   'False
         Left            =   4080
         Top             =   180
      End
      Begin MSMask.MaskEdBox mskTimerSeconds 
         Height          =   315
         Left            =   2880
         TabIndex        =   1
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
         Left            =   7440
         TabIndex        =   2
         Top             =   240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFileNameEUT 
         Height          =   255
         Left            =   7440
         TabIndex        =   3
         Top             =   480
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt_imppos 
         Height          =   255
         Left            =   7440
         TabIndex        =   14
         Top             =   720
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt_hispos 
         Height          =   255
         Left            =   7440
         TabIndex        =   15
         Top             =   960
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdUnidad 
         Height          =   255
         Left            =   1440
         TabIndex        =   18
         Top             =   720
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskRepos 
         Height          =   255
         Left            =   3600
         TabIndex        =   20
         Top             =   1200
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox DIR_CLICKFIND 
         Height          =   255
         Left            =   7440
         TabIndex        =   22
         Top             =   1200
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Directorio Recep Ficheros GPS"
         Height          =   255
         Index           =   4
         Left            =   4920
         TabIndex        =   23
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Texto1pos 
         Caption         =   "Rellenar Posiciones 0 Mins = No Rellenar"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   21
         Top             =   1200
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "Unidad Red"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   19
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Directorio Historico Posiciones"
         Height          =   255
         Index           =   2
         Left            =   4920
         TabIndex        =   17
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Directorio Importacion Posiciones"
         Height          =   255
         Index           =   1
         Left            =   4920
         TabIndex        =   16
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label lblMinuti 
         Alignment       =   1  'Right Justify
         Caption         =   "Esecuzione delle procedure ogni                  secondi"
         Height          =   495
         Left            =   180
         TabIndex        =   6
         Top             =   240
         Width           =   3915
      End
      Begin VB.Label Label1 
         Caption         =   "Directorio GIS TPH"
         Height          =   255
         Index           =   14
         Left            =   5880
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Directorio GIS EUT:"
         Height          =   255
         Index           =   0
         Left            =   5880
         TabIndex        =   4
         Top             =   480
         Width           =   1575
      End
   End
   Begin TrueDBGrid60.TDBGrid gridProcedures 
      Height          =   2415
      Left            =   0
      OleObjectBlob   =   "frmSatellitarePos.frx":0408
      TabIndex        =   13
      Top             =   2280
      Width           =   11175
   End
End
Attribute VB_Name = "frmSatellitarePos"
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
Private Const REGISTRY_KEY_NAME_PATH_POSOES = "PATH_POSOES"
Private Const REGISTRY_KEY_NAME_PATH_HISPOES = "PATH_HISPOES"
Private Const REGISTRY_KEY_NAME_PATH_CLICKFIND = "PATH_CLICKFIND"
Private Const REGISTRY_KEY_NAME_UNIDAD = "Unidadresoil"
Private Const REGISTRY_KEY_NAME_REPOS = "Rellenar_Posiciones"
Private mobjOutFileTPH As CTextFile
Private mobjOutFileEUT As CTextFile
Private mobjOutFileCLICK As CTextFile
Private mobjOutFileSTARTIT As CTextFile
Private mobjOutFileSaturacionTot As CTextFile
Private mobjOutFileSaturacionTram As CTextFile
Private mobjOutFile As CTextFile
Private ProcesoEnCurso As Boolean
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

        XA.value(2, 1) = "In esecuzione."
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
    
    XA.value(2, 1) = "In attesa."
    gridProcedures.RefetchCol 1

    Execute_Trasm_Posizioni_CEPSA = True
    Set KVE = Nothing
    Exit Function

ErrHandler:
    Set KVE = Nothing
    mobjOutFile.CloseFile
    mobjOutFileEUT.CloseFile
    mobjOutFileTPH.CloseFile
    'XA.value(2, 1) = "In errore."
    gridProcedures.RefetchCol 1
End Function

Public Function GeneraSaturacion()
If (ProcesoEnCurso = False) Then
ProcesoEnCurso = True
End If

Dim VelocidadMedia, TiempoCarga, TiempoDescarga As Integer
Dim Tiemporod, TiempoAct As Double
Dim DA As CDataAccess
Set DA = DataConnect.DA
LogMessage "Inicio proceso de Generación Saturacion" & vbCrLf
XA.value(0, 1) = "En Ejecucion"
gridProcedures.RefetchCol 1
Dim FechaCarga, HoraCarga, MesFi, DiaFi, FechaPLani As String
Dim VehiculoAsignado As String
Dim Motivo As String
Dim SQLOrdenes As String
Dim SQLInsertEntrega As String
Dim varKBufferOrdenes As Variant
Dim lngHowManyOrdenes As Long
Dim SQLBuscaCP As String
Dim varKBufferCliCP As Variant
Dim lngHowManyCliCP As Long
Dim SQLSatVehiculo As String
Dim varKBufferSatvehiculo As Variant
Dim lngHowManySatVehiculo As Long
Dim o, i As Integer
Dim NOrden, Depoorden, FechaOrden As String
Dim CantOrden As Long
Dim Nviaje As Integer
Dim Nentrega As Integer
Dim Distancia As Integer
Dim DistanciaRad As Double
Dim Tiempo As Integer
Dim Longx As String
Dim Laty As String
Dim LongXAnt As String
Dim LatYAnt As String
Dim LongxDep As String
Dim LatYDep As String
Dim CodPos As String
Dim Codcli As String
Dim Horaini As Date
Dim HoraLlegada As Date
Dim HoraEntrega As Date
Dim MatTr As String
Dim MatCis As String
Dim HoraFin As Date
Dim HorainiCalc As Date
Dim HoraFinCalc As Date
Dim HoraFinRea As Date
Dim TiempoOcupado, TiempoDisponible As Integer
Dim TiempoTramo As Integer
Dim IndiceOcupa As Double
Dim Condicionante As String
Dim Bomba As Boolean
Dim TanqueDedicado As String
Dim ret As Boolean
Dim xy As RESOIL2.CTextFile
Set xy = New CTextFile
Dim SQLInsertUTSaturacinoAuto As String

SQLOrdenes = "SELECT        ORDIOCT0.OCTIDOCT, ORDIOCT0.OCTIDASC, ORDIOCT0.OCTIDADZ, ORDIOCT0.OCTIDADE, ORDIOCT0.OCTDTCRI, SUM(ORDIOCR0.OCRQRILT) AS Litros, ORDIOCT0.OCTHRFR1, ORDIOCT0.OCTHRTO1, " & _
"                         ANAGADZ0.ADZCGPSX, ANAGADZ0.ADZCGPSY, ANAGADE0.ADECGPSX, ANAGADE0.ADECGPSY, ANAGADZ0.ADZNRCAP, ANAGADZ0.ADZLOCAL, ANAGDED0.DEDKMATT, ORDIOCT0.OCTRIFE1, " & _
"                         ORDIOCT0.OCTCONTE, SATURACION.SAROCT, ORDIOCT0.OCTFLSCA " & _
"FROM            ORDIOCT0 INNER JOIN " & _
"                         ORDIOCR0 ON ORDIOCT0.OCTIDOCT = ORDIOCR0.OCRIDOCT INNER JOIN " & _
"                         ANAGADZ0 ON ORDIOCT0.OCTIDADZ = ANAGADZ0.ADZIDADZ INNER JOIN " & _
"                         ANAGADE0 ON ORDIOCT0.OCTIDADE = ANAGADE0.ADEIDADE LEFT OUTER JOIN " & _
"                         ANAGDED0 ON ORDIOCT0.OCTIDADE = ANAGDED0.DEDIDADE AND ORDIOCT0.OCTIDADZ = ANAGDED0.DEDIDADZ LEFT OUTER JOIN " & _
"                         SATURACION ON ORDIOCT0.OCTIDOCT = SATURACION.SAROCT " & _
"GROUP BY ORDIOCT0.OCTIDASC, ORDIOCT0.OCTIDADE, ORDIOCT0.OCTDTCRI, ORDIOCT0.OCTIDOCT, ORDIOCT0.OCTIDADZ, ORDIOCT0.OCTHRFR1, ORDIOCT0.OCTHRTO1, ANAGADZ0.ADZCGPSX, ANAGADZ0.ADZCGPSY, " & _
"                         ANAGADZ0.ADZNRCAP, ANAGDED0.DEDKMATT, ANAGADE0.ADECGPSX, ANAGADE0.ADECGPSY, ANAGADZ0.ADZLOCAL, ORDIOCT0.OCTRIFE1, ORDIOCT0.OCTCONTE, SATURACION.SAROCT, ORDIOCT0.OCTFLSCA " & _
"HAVING        (ORDIOCT0.OCTIDASC = '038') AND (ORDIOCT0.OCTIDADE = '004' OR " & _
"                         ORDIOCT0.OCTIDADE = '005' OR ORDIOCT0.OCTIDADE = '021' OR " & _
"                         ORDIOCT0.OCTIDADE = '008' OR " & _
"                         ORDIOCT0.OCTIDADE = '010' OR " & _
"                         ORDIOCT0.OCTIDADE = '002' OR " & _
"                         ORDIOCT0.OCTIDADE = '028' OR " & _
"                         ORDIOCT0.OCTIDADE = '016' OR " & _
"                         ORDIOCT0.OCTIDADE = '030' OR " & _
"                         ORDIOCT0.OCTIDADE = '031' OR " & _
"                         ORDIOCT0.OCTIDADE = '032' OR " & _
"                         ORDIOCT0.OCTIDADE = '033' OR " & _
"                         ORDIOCT0.OCTIDADE = '034' OR " & _
"                         ORDIOCT0.OCTIDADE = '035') AND (ORDIOCT0.OCTDTCRI > GETDATE() -1) AND (SATURACION.SAROCT IS NULL) AND (ORDIOCT0.OCTFLSCA = 'N') AND (SUM(ORDIOCR0.OCRQRILT) < 9500) " & _
"ORDER BY ORDIOCT0.OCTDTCRI,ORDIOCT0.OCTIDADE,  ANAGADZ0.ADZNRCAP"

' añadir AND (SUM(ORDIOCR0.OCRQRILT) < 9500) para filiales
' fecha exacta CONVERT(DATETIME, '2015-09-11 00:00:00', 102)
' Borra planificacion


' Borramos todo si ponemos 0, si ponemos 1 borramos solo los depositos que tiene alguna orden sin planificar
If SatBorraPLanif(1) Then
End If

If DA.QueryCommand(SQLOrdenes, varKBufferOrdenes, lngHowManyOrdenes) Then
End If


If lngHowManyOrdenes > 0 Then

    For o = 0 To UBound(varKBufferOrdenes, 2)
    
    NOrden = varKBufferOrdenes(0, o)
    FechaOrden = varKBufferOrdenes(4, o)
    FechaOrden = Format(FechaOrden, "yyyy-mm-dd hh:mm:ss")
    Depoorden = varKBufferOrdenes(3, o)
    CantOrden = varKBufferOrdenes(5, o)
    Codcli = varKBufferOrdenes(2, o)
    LongxDep = varKBufferOrdenes(10, o)
    LatYDep = varKBufferOrdenes(11, o)
    Condicionante = varKBufferOrdenes(16, o)
    TanqueDedicado = varKBufferOrdenes(18, o)
    
    Dim strCond As String
   If Mid(Condicionante, 4, 1) > "0" Then
   Bomba = True
   Else
   Bomba = False
   End If

    
    If IsNull(varKBufferOrdenes(8, o)) Then
        Longx = ""
    Else
        Longx = varKBufferOrdenes(8, o)
    End If
    If IsNull(varKBufferOrdenes(9, o)) Then
        Laty = ""
    Else
        Laty = varKBufferOrdenes(9, o)
    End If


    If SatBuscaVehiculo(FechaOrden, Depoorden, NOrden, CantOrden, VehiculoAsignado, Nviaje, Nentrega, Motivo, LongXAnt, LatYAnt, Horaini, HoraLlegada, Bomba, TanqueDedicado) Then
    End If
    If IsNull(varKBufferOrdenes(12, o)) Then
        CodPos = ""
    Else
        CodPos = varKBufferOrdenes(12, o)
    End If

    
    'Si no sabemos la longitud y latitud la cogemos otro cliente por el código postal
    If Longx = "" Or Laty = "" Or Longx = "0" Or Laty = "0" Then

    SQLBuscaCP = "SELECT ADZCGPSX, ADZCGPSY FROM ANAGADZ0 WHERE (ADZCGPSX <> '' AND ADZCGPSX <> '0') AND (ADZNRCAP = '" & CodPos & "')"
    If DA.QueryCommand(SQLBuscaCP, varKBufferCliCP, lngHowManyCliCP) Then
    End If
    
    If lngHowManyCliCP > 0 Then
 
    
    Longx = CStr(varKBufferCliCP(0, 0))
    Laty = CStr(varKBufferCliCP(1, 0))
    Else
    Motivo = "No hay ninguna posición por el CP " & CodPos & " Cli:" & Codcli
    End If
    End If
    
   ' Calculamos la distancia
    If Nentrega = 1 Then
    LongXAnt = LongxDep
    LatYAnt = LatYDep
    End If
    
       
    If LongXAnt <> "" And LatYAnt <> "" And Longx <> "" And Laty <> "" Then
    DistanciaRad = SAtCalcDistCoordenadas(CDec(LongXAnt), CDec(LatYAnt), CDec(Longx), CDec(Laty))
    Distancia = CInt(DistanciaRad) * 1.4
    Else: Distancia = 0
    End If
    
    'Calculamos el tiempo
    VelocidadMedia = 30
    TiempoCarga = 20
    TiempoDescarga = 20
    Tiemporod = Distancia / 48 * 60
    If Nentrega = 1 Then
    TiempoAct = TiempoCarga + TiempoDescarga
    Else
    TiempoAct = TiempoDescarga
    End If
    
    Tiempo = TiempoAct + Tiemporod
    
    'Calculamos la hora de llegada
    If Nentrega = 1 Then
     HoraEntrega = DateAdd("n", Tiempo, Horaini)
    Else
    HoraEntrega = DateAdd("n", Tiempo, HoraLlegada)
    End If
    HoraEntrega = Format(HoraEntrega, "yyyy-mm-dd hh:mm:ss")
    
    If Nviaje > 1 And Nentrega = 1 Then HoraEntrega = DateAdd("n", Tiempo, HoraLlegada)
    
    ' Updateamos la tabla Tanques
    If SatCHKTancaje(FechaOrden, NOrden, VehiculoAsignado, Nviaje, TanqueDedicado, True) Then
    End If
    
    'insertamos la entrega
    SQLInsertEntrega = "INSERT INTO SATURACION(SARFECHA, SARUT, SAROCT, SARTIEMPO, SARDIST, SARNVI, SARNDESC, SAROBSERVA, SARLITROS, SARLONGX, SARLATY, SARHLLEGADA, SARADE)" & _
    "VALUES        (CONVERT(DATETIME,'" & FechaOrden & "' , 102),'" & VehiculoAsignado & "','" & NOrden & "','" & Tiempo & "','" & Distancia & "','" & Nviaje & "','" & Nentrega & " ','" & Motivo & "'," & CantOrden & ",'" & Longx & "','" & Laty & "',CONVERT(DATETIME, '" & HoraEntrega & "', 102),'" & Depoorden & "')"

    If Not DA.ExecuteQuery(SQLInsertEntrega) Then
    End If
    Next o
End If

FechaCarga = Format$(Date, "YYYYMMDD")
HoraCarga = Format$(Time, "HHMM")
MesFi = Format$(Date, "MM")
DiaFi = Format$(Date, "DD")

' Generamos el fichero de Saturación total

'If mobjOutFileSaturacionTot.OpenFile(mskFileNameEUT.Text + "SAT1_" & MesFi & DiaFi & "_" & FechaCarga & HoraCarga + ".DAT", 1) Then
'End If
'
''            For i = 0 To UBound(varKBufferTPH, 2)
'                mobjOutFileSaturacionTot.Item("1").value = FechaCarga & HoraCarga
''                mobjOutFileSaturacionTot.Item("2").value = NLoteTPH
'                mobjOutFileSaturacionTot.Item("3").value = "0004"
''                mobjOutFileSaturacionTot.Item("4").value = Format$(Time, "hh:mm:ss")
''                mobjOutFileSaturacionTot.Item("5").value = lngHowManyTPH
''                mobjOutFileSaturacionTot.Item("6").value = varKBufferTPH(2, i)
''                mobjOutFileSaturacionTot.Item("7").value = varKBufferTPH(3, i)
'                mobjOutFileSaturacionTot.WriteLine
''            Next i
'
'mobjOutFileSaturacionTot.CloseFile

' Generamos el fichero de Saturación por tramos

'If mobjOutFileSaturacionTram.OpenFile(mskFileNameEUT.Text + "SAT2_" & MesFi & DiaFi & "_" & FechaCarga & HoraCarga + ".DAT", 1) Then
'End If
'
''            For i = 0 To UBound(varKBufferTPH, 2)
'                mobjOutFileSaturacionTram.Item("1").value = FechaCarga & HoraCarga
''                mobjOutFileSaturacionTram.Item("2").value = NLoteTPH
'                mobjOutFileSaturacionTram.Item("3").value = "0004"
''                mobjOutFileSaturacionTram.Item("4").value = Format$(Time, "hh:mm:ss")
''                mobjOutFileSaturacionTram.Item("5").value = lngHowManyTPH
''                mobjOutFileSaturacionTram.Item("6").value = varKBufferTPH(2, i)
''                mobjOutFileSaturacionTram.Item("7").value = varKBufferTPH(3, i)
''                mobjOutFileSaturacionTram.Item("8").value = varKBufferTPH(3, i)
''                mobjOutFileSaturacionTram.Item("9").value = varKBufferTPH(3, i)
''                mobjOutFileSaturacionTram.Item("10").value = varKBufferTPH(3, i)
''                mobjOutFileSaturacionTram.Item("11").value = varKBufferTPH(3, i)
'                mobjOutFileSaturacionTram.WriteLine
''            Next i
'
'mobjOutFileSaturacionTram.CloseFile


' Generamos el fichero de Saturación por vehiculo
Dim Vehimodi As Boolean
Vehimodi = True

SQLSatVehiculo = "SELECT        SATURADISP.SATFECHA, ANAGAUT0.AUTIDADE, MULTSDE0.SDECODSO AS CODDEP, SATURADISP.SATUT, MULTSUT0.SUTCODSO, ANAGAUT0.AUTIDAM1, CASE WHEN ANAGAUT0.AUTIDAM2 <> '' OR " & _
"                         ANAGAUT0.AUTIDAM2 <> NULL THEN ANAGAUT0.AUTIDAM2 ELSE ANAGAUT0.AUTIDAM1 END AS Matcis, SATURADISP.SATHINI, CONVERT(datetime, DATEADD(hh, SATURADISP.SATUSO, SATURADISP.SATHINI), " & _
"                         102) AS horafin, MAX(SATURACION.SARHLLEGADA) AS HFINREA, CASE WHEN MAX(SATURACION.SARHLLEGADA) > CONVERT(datetime, DATEADD(hh, SATURADISP.SATUSO, SATURADISP.SATHINI), 102) " & _
"                         THEN MAX(SATURACION.SARHLLEGADA) ELSE CONVERT(datetime, DATEADD(hh, SATURADISP.SATUSO, SATURADISP.SATHINI), 102) END AS HFINCALC " & _
"FROM            SATURADISP LEFT OUTER JOIN" & _
"                         SATURACION ON SATURADISP.SATFECHA = SATURACION.SARFECHA AND SATURADISP.SATUT = SATURACION.SARUT LEFT OUTER JOIN " & _
"                         MULTSDE0 RIGHT OUTER JOIN " & _
"                         ANAGADE0 ON MULTSDE0.SDEIDADE = ANAGADE0.ADEIDADE RIGHT OUTER JOIN " & _
"                         MULTSUT0 RIGHT OUTER JOIN " & _
"                         ANAGADE0 AS ANAGADE0_1 INNER JOIN " & _
"                         ANAGAUT0 ON ANAGADE0_1.ADEIDADE = ANAGAUT0.AUTIDADE ON MULTSUT0.SUTIDASC = ANAGADE0_1.ADEIDASC AND MULTSUT0.SUTIDAUT = ANAGAUT0.AUTIDAUT ON " & _
"                         ANAGADE0.ADEIDADE = ANAGAUT0.AUTIDADE ON SATURADISP.SATUT = ANAGAUT0.AUTIDAUT " & _
"GROUP BY ANAGAUT0.AUTIDADE, SATURADISP.SATFECHA, MULTSDE0.SDECODSO, MULTSDE0.SDEIDASC, MULTSUT0.SUTCODSO, SATURADISP.SATUT, ANAGAUT0.AUTIDAM1, ANAGAUT0.AUTIDAM2, " & _
"                         SATURADISP.SATHINI, SATURADISP.SATUSO, SATURADISP.SATDENV " & _
"HAVING        (MULTSDE0.SDEIDASC = '901') AND (SATURADISP.SATFECHA > GETDATE() - 1) AND (SATURADISP.SATDENV = 0) "

If DA.QueryCommand(SQLSatVehiculo, varKBufferSatvehiculo, lngHowManySatVehiculo) Then
End If


If lngHowManySatVehiculo <= 0 Then

        SQLSatVehiculo = "SELECT        SATURADISP.SATFECHA, ANAGAUT0.AUTIDADE, MULTSDE0.SDECODSO AS CODDEP, SATURADISP.SATUT, MULTSUT0.SUTCODSO, ANAGAUT0.AUTIDAM1, CASE WHEN ANAGAUT0.AUTIDAM2 <> '' OR " & _
"                         ANAGAUT0.AUTIDAM2 <> NULL THEN ANAGAUT0.AUTIDAM2 ELSE ANAGAUT0.AUTIDAM1 END AS Matcis, SATURADISP.SATHINI, CONVERT(datetime, DATEADD(hh, SATURADISP.SATUSO, SATURADISP.SATHINI), " & _
"                         102) AS horafin, MAX(SATURACION.SARHLLEGADA) AS HFINREA, CASE WHEN MAX(SATURACION.SARHLLEGADA) > CONVERT(datetime, DATEADD(hh, SATURADISP.SATUSO, SATURADISP.SATHINI), 102) " & _
"                         THEN MAX(SATURACION.SARHLLEGADA) ELSE CONVERT(datetime, DATEADD(hh, SATURADISP.SATUSO, SATURADISP.SATHINI), 102) END AS HFINCALC " & _
"FROM            SATURADISP LEFT OUTER JOIN " & _
"                         SATURACION ON SATURADISP.SATFECHA = SATURACION.SARFECHA AND SATURADISP.SATUT = SATURACION.SARUT LEFT OUTER JOIN " & _
"                         MULTSDE0 RIGHT OUTER JOIN " & _
"                         ANAGADE0 ON MULTSDE0.SDEIDADE = ANAGADE0.ADEIDADE RIGHT OUTER JOIN " & _
"                         MULTSUT0 RIGHT OUTER JOIN " & _
"                         ANAGADE0 AS ANAGADE0_1 INNER JOIN " & _
"                         ANAGAUT0 ON ANAGADE0_1.ADEIDADE = ANAGAUT0.AUTIDADE ON MULTSUT0.SUTIDASC = ANAGADE0_1.ADEIDASC AND MULTSUT0.SUTIDAUT = ANAGAUT0.AUTIDAUT ON " & _
"                         ANAGADE0.ADEIDADE = ANAGAUT0.AUTIDADE ON SATURADISP.SATUT = ANAGAUT0.AUTIDAUT " & _
"GROUP BY ANAGAUT0.AUTIDADE, SATURADISP.SATFECHA, MULTSDE0.SDECODSO, MULTSDE0.SDEIDASC, MULTSUT0.SUTCODSO, SATURADISP.SATUT, ANAGAUT0.AUTIDAM1, ANAGAUT0.AUTIDAM2, " & _
"                         SATURADISP.SATHINI, SATURADISP.SATUSO, SATURACION.SATENV " & _
"HAVING        (MULTSDE0.SDEIDASC = '901') AND (SATURADISP.SATFECHA > GETDATE() - 1) AND (SATURACION.SATENV = 0) "
        Vehimodi = False
End If

If DA.QueryCommand(SQLSatVehiculo, varKBufferSatvehiculo, lngHowManySatVehiculo) Then
End If

If lngHowManySatVehiculo > 0 Then

            If mobjOutFileSaturacionTram.OpenFile(mskFileNameEUT.Text + "SAT3_" & MesFi & DiaFi & "_" & FechaCarga & HoraCarga + ".DAT", 1) Then
            End If
            ' Calculamos el tramo 00 hasta las 8:00 de la mañana.
            For i = 0 To UBound(varKBufferSatvehiculo, 2)
                mobjOutFileSaturacionTram.Item("1").value = FechaCarga & HoraCarga
                FechaPLani = varKBufferSatvehiculo(0, i)
                FechaPLani = Format$(FechaPLani, "YYYYMMDD")
                mobjOutFileSaturacionTram.Item("2").value = FechaPLani
                mobjOutFileSaturacionTram.Item("3").value = "3369" 'Oilogist 0004 Eutrasur
                mobjOutFileSaturacionTram.Item("4").value = varKBufferSatvehiculo(2, i)
                mobjOutFileSaturacionTram.Item("5").value = RellenaEsp("0", 2)
                mobjOutFileSaturacionTram.Item("6").value = RellenaEsp(varKBufferSatvehiculo(4, i), 11)
                MatTr = varKBufferSatvehiculo(5, i)
                MatTr = Replace(MatTr, " ", "")
                MatTr = Replace(MatTr, "-", "")
                mobjOutFileSaturacionTram.Item("7").value = RellenaEsp(MatTr, 18)
                MatCis = varKBufferSatvehiculo(6, i)
                MatCis = Replace(MatCis, " ", "")
                MatCis = Replace(MatCis, "-", "")
                mobjOutFileSaturacionTram.Item("8").value = RellenaEsp(MatCis, 18)
                
                
                
                ' Calculamos la disponibilidad
                
                TiempoDisponible = 0
                HoraFin = varKBufferSatvehiculo(8, i)
                HorainiCalc = varKBufferSatvehiculo(7, i)
                HoraFinCalc = varKBufferSatvehiculo(10, i)
                If Not IsNull(varKBufferSatvehiculo(9, i)) Then
                HoraFinRea = varKBufferSatvehiculo(9, i)
                Else
                HoraFinRea = #1/1/1900#
                End If
                
                
                If HoraFinRea > #1/1/1900 8:00:00 AM# Then
                HoraFinRea = #1/1/1900 8:00:00 AM#
                End If
                TiempoOcupado = DateDiff("n", HorainiCalc, HoraFinRea)
                If TiempoOcupado < 0 Then TiempoOcupado = 0
                
                
                
                
                If HoraFinCalc > #1/1/1900 8:00:00 AM# Then
                HoraFinCalc = #1/1/1900 8:00:00 AM#
                End If
                TiempoTramo = DateDiff("n", HorainiCalc, HoraFinCalc)
                
                
                TiempoDisponible = TiempoTramo - TiempoOcupado
                If TiempoDisponible < 0 Then TiempoDisponible = 0

                                
                mobjOutFileSaturacionTram.Item("9").value = RellenaEsp(TiempoOcupado, 6)
                mobjOutFileSaturacionTram.Item("10").value = RellenaEsp(TiempoDisponible, 6)
                
                ' Calculamos el indice de ocupación
                If TiempoDisponible > 0 Then
                IndiceOcupa = TiempoOcupado / TiempoDisponible
                Else
                IndiceOcupa = 0
                End If
                mobjOutFileSaturacionTram.Item("11").value = RellenaEsp(Replace(Format(IndiceOcupa, "0.00000000"), ",", "."), 12)
                If mobjOutFileSaturacionTram.WriteLine(, True) Then
                End If
            Next i
            
            
            ' Calculamos el tramo 01 desde las 8:00 hasta las 14:00 de la mañana.
            For i = 0 To UBound(varKBufferSatvehiculo, 2)
                mobjOutFileSaturacionTram.Item("1").value = FechaCarga & HoraCarga
                FechaPLani = varKBufferSatvehiculo(0, i)
                FechaPLani = Format$(FechaPLani, "YYYYMMDD")
                mobjOutFileSaturacionTram.Item("2").value = FechaPLani
                mobjOutFileSaturacionTram.Item("3").value = "3369" 'Oilogist 0004 Eutrasur
                mobjOutFileSaturacionTram.Item("4").value = varKBufferSatvehiculo(2, i)
                mobjOutFileSaturacionTram.Item("5").value = RellenaEsp("1", 2)
                mobjOutFileSaturacionTram.Item("6").value = RellenaEsp(varKBufferSatvehiculo(4, i), 11)
                MatTr = varKBufferSatvehiculo(5, i)
                MatTr = Replace(MatTr, " ", "")
                MatTr = Replace(MatTr, "-", "")
                mobjOutFileSaturacionTram.Item("7").value = RellenaEsp(MatTr, 18)
                MatCis = varKBufferSatvehiculo(6, i)
                MatCis = Replace(MatCis, " ", "")
                MatCis = Replace(MatCis, "-", "")
                mobjOutFileSaturacionTram.Item("8").value = RellenaEsp(MatCis, 18)
                
                ' Calculamos la disponibilidad
                
                TiempoDisponible = 0
                HoraFin = varKBufferSatvehiculo(8, i)
                HorainiCalc = varKBufferSatvehiculo(7, i)
                HoraFinCalc = varKBufferSatvehiculo(10, i)
                If Not IsNull(varKBufferSatvehiculo(9, i)) Then
                HoraFinRea = varKBufferSatvehiculo(9, i)
                Else
                HoraFinRea = #1/1/1900#
                End If
                
                
                If HoraFinRea > #1/1/1900 2:00:00 PM# Then
                HoraFinRea = #1/1/1900 2:00:00 PM#
                End If
                
                If HorainiCalc < #1/1/1900 8:00:00 AM# Then
                HorainiCalc = #1/1/1900 8:00:00 AM#
                End If
                
                TiempoOcupado = DateDiff("n", HorainiCalc, HoraFinRea)
                If TiempoOcupado < 0 Then TiempoOcupado = 0
                
                               
                
                If HoraFinCalc > #1/1/1900 2:00:00 PM# Then
                HoraFinCalc = #1/1/1900 2:00:00 PM#
                End If
                TiempoTramo = DateDiff("n", HorainiCalc, HoraFinCalc)
                
                
                TiempoDisponible = TiempoTramo - TiempoOcupado
                If TiempoDisponible < 0 Then TiempoDisponible = 0
                                
                mobjOutFileSaturacionTram.Item("9").value = RellenaEsp(TiempoOcupado, 6)
                mobjOutFileSaturacionTram.Item("10").value = RellenaEsp(TiempoDisponible, 6)
                
                ' Calculamos el indice de ocupación
                If TiempoDisponible > 0 Then
                IndiceOcupa = TiempoOcupado / TiempoDisponible
                Else
                IndiceOcupa = 0
                End If
                mobjOutFileSaturacionTram.Item("11").value = RellenaEsp(Replace(Format(IndiceOcupa, "0.00000000"), ",", "."), 12)
                If mobjOutFileSaturacionTram.WriteLine(, True) Then
                End If
            Next i

            ' Calculamos el tramo 02 desde las 14:00 a las 20:00
            For i = 0 To UBound(varKBufferSatvehiculo, 2)
                mobjOutFileSaturacionTram.Item("1").value = FechaCarga & HoraCarga
                FechaPLani = varKBufferSatvehiculo(0, i)
                FechaPLani = Format$(FechaPLani, "YYYYMMDD")
                mobjOutFileSaturacionTram.Item("2").value = FechaPLani
                mobjOutFileSaturacionTram.Item("3").value = "3369" 'Oilogist 0004 Eutrasur
                mobjOutFileSaturacionTram.Item("4").value = varKBufferSatvehiculo(2, i)
                mobjOutFileSaturacionTram.Item("5").value = RellenaEsp("2", 2) ' TRamo
                mobjOutFileSaturacionTram.Item("6").value = RellenaEsp(varKBufferSatvehiculo(4, i), 11)
                MatTr = varKBufferSatvehiculo(5, i)
                MatTr = Replace(MatTr, " ", "")
                MatTr = Replace(MatTr, "-", "")
                mobjOutFileSaturacionTram.Item("7").value = RellenaEsp(MatTr, 18)
                MatCis = varKBufferSatvehiculo(6, i)
                MatCis = Replace(MatCis, " ", "")
                MatCis = Replace(MatCis, "-", "")
                mobjOutFileSaturacionTram.Item("8").value = RellenaEsp(MatCis, 18)
                
                

                
                ' Calculamos la disponibilidad

                
                TiempoDisponible = 0
                HoraFin = varKBufferSatvehiculo(8, i)
                HorainiCalc = varKBufferSatvehiculo(7, i)
                HoraFinCalc = varKBufferSatvehiculo(10, i)
                If Not IsNull(varKBufferSatvehiculo(9, i)) Then
                HoraFinRea = varKBufferSatvehiculo(9, i)
                Else
                HoraFinRea = #1/1/1900#
                End If
                
                
                If HoraFinRea > #1/1/1900 8:00:00 PM# Then
                HoraFinRea = #1/1/1900 8:00:00 PM#
                End If
                
                If HorainiCalc < #1/1/1900 2:00:00 PM# Then
                HorainiCalc = #1/1/1900 2:00:00 PM#
                End If
                
                TiempoOcupado = DateDiff("n", HorainiCalc, HoraFinRea)
                If TiempoOcupado < 0 Then TiempoOcupado = 0
                
                               
                
                If HoraFinCalc > #1/1/1900 8:00:00 PM# Then
                HoraFinCalc = #1/1/1900 8:00:00 PM#
                End If
                TiempoTramo = DateDiff("n", HorainiCalc, HoraFinCalc)
                
                
                TiempoDisponible = TiempoTramo - TiempoOcupado
                If TiempoDisponible < 0 Then TiempoDisponible = 0
                                
                
                mobjOutFileSaturacionTram.Item("9").value = RellenaEsp(TiempoOcupado, 6)
                mobjOutFileSaturacionTram.Item("10").value = RellenaEsp(TiempoDisponible, 6)
                
                ' Calculamos el indice de ocupación
                If TiempoDisponible > 0 Then
                IndiceOcupa = TiempoOcupado / TiempoDisponible
                Else
                IndiceOcupa = 0
                End If
                mobjOutFileSaturacionTram.Item("11").value = RellenaEsp(Replace(Format(IndiceOcupa, "0.00000000"), ",", "."), 12)
                If mobjOutFileSaturacionTram.WriteLine(, True) Then
                End If
            Next i


        mobjOutFileSaturacionTram.CloseFile
        ' Marcamos todo como enviado
        
        If Vehimodi = False Then
            If Not DA.ExecuteQuery("UPDATE SATURACION SET SATENV = 1 WHERE (SATENV = 0)") Then
            End If
        Else
            If Not DA.ExecuteQuery("UPDATE SATURADISP SET SATDENV = 1 WHERE (SATDENV = 0)") Then
            End If
        End If
    Else
    ' Aqui podemos poner controles de no poner vehículos.
    ret = xy.OpenFile(MaskEdUnidad.Text & ":\Resoil2\Alarmas\SINSATURACION.TXT", 1)
    ret = xy.CloseFile()
    
    ' Si queremos insertamos automaticamente las UT del dia anterior.
'    SQLInsertUTSaturacinoAuto = ""
'    If Not DA.ExecuteQuery(SQLInsertUTSaturacinoAuto) Then
'    End If
    
    
    End If

LogMessage "Fin proceso de Generación Saturacion" & vbCrLf
XA.value(0, 1) = "En Espera"
gridProcedures.RefetchCol 1
ProcesoEnCurso = False

End Function
Public Function Captura_prec(ByVal MatCisterna As String, ByVal FechaEv As String, ByVal Latitud As String, ByVal Longitud As String, ByVal Speed As Integer, ByVal Dep1 As String, ByVal Dep2 As String, ByVal Dep3 As String, ByVal Dep4 As String, ByVal Dep5 As String, ByVal Dep6 As String) As Boolean

Dim Varkbuffer As Variant
Dim VarkbufferEC0 As Variant
Dim VarkbufferEDep As Variant
Dim VarkbufferPOS As Variant
Dim lngHowMany As Long
Dim lngHowManyEC0 As Long
Dim lngHowManyEDep As Long
Dim lngHowManyPOS As Long
Dim DA As CDataAccess
Dim RegistroLeid, MatCisternaTh, FechaEvento, Deposito, EstadoDep, Latitudgr, Longitudgr, Mattract, CodGPS, FechaPos As String
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
Dim Signolat, Signolong, LongitudCampo As Integer
Dim SQL_INSERTA_PREC As String
On Error GoTo ErrHandler

SQL_ULT_EST_PREC = "SELECT     dbo.ANAGAAM0.AAMIDPRE, dbo.ANAGAEP0.AEPESTHM AS DP1, ANAGAEP0_1.AEPESTHM AS DP2, ANAGAEP0_2.AEPESTHM AS DP3, " & _
"                      ANAGAEP0_3.AEPESTHM AS DP4, ANAGAEP0_4.AEPESTHM AS DP5, ANAGAEP0_5.AEPESTHM AS DP6, " & _
"dbo.ANAGAEC0.AECEPDP1,dbo.ANAGAEC0.AECEPDP2, dbo.ANAGAEC0.AECEPDP3, dbo.ANAGAEC0.AECEPDP4, dbo.ANAGAEC0.AECEPDP5, dbo.ANAGAEC0.AECEPDP6 " & _
"FROM         dbo.ANAGAEC0 LEFT OUTER JOIN  dbo.ANAGAEP0 ON dbo.ANAGAEC0.AECEPDP1 = dbo.ANAGAEP0.AEPESCEP LEFT OUTER JOIN " & _
"                      dbo.ANAGAEP0 ANAGAEP0_5 ON dbo.ANAGAEC0.AECEPDP6 = ANAGAEP0_5.AEPESCEP LEFT OUTER JOIN " & _
"                      dbo.ANAGAEP0 ANAGAEP0_4 ON dbo.ANAGAEC0.AECEPDP5 = ANAGAEP0_4.AEPESCEP LEFT OUTER JOIN " & _
"                      dbo.ANAGAEP0 ANAGAEP0_3 ON dbo.ANAGAEC0.AECEPDP4 = ANAGAEP0_3.AEPESCEP LEFT OUTER JOIN " & _
"                      dbo.ANAGAEP0 ANAGAEP0_2 ON dbo.ANAGAEC0.AECEPDP3 = ANAGAEP0_2.AEPESCEP LEFT OUTER JOIN " & _
"                      dbo.ANAGAEP0 ANAGAEP0_1 ON dbo.ANAGAEC0.AECEPDP2 = ANAGAEP0_1.AEPESCEP LEFT OUTER JOIN " & _
"                      dbo.ANAGAAM0 ON dbo.ANAGAEC0.AECIDAAM = dbo.ANAGAAM0.AAMIDAAM WHERE     (dbo.ANAGAAM0.AAMIDAAM = ?)"
  
    
    
    ' buscamos cual era el ultimo estado de la cisterna en nuestra bd
    If Not DA.QueryCommand(SQL_ULT_EST_PREC, VarkbufferEC0, lngHowManyEC0, MatCisterna) Then
            LogMessage "Error en el proceso de captura de precintos Thaumat. " & vbCrLf
            gridProcedures.RefetchCol 1
    End If
    ' si la cisterna no estaba en nuestro histórico la damos de alta y volvemos a mirar los estados aunque esten a vacio para seguir
    If lngHowManyEC0 = 0 And MatCisterna <> "" Then
        If Not (DA.ExecuteQuery("INSERT INTO ANAGAEC0 (AECIDAAM) VALUES('" & MatCisterna & "')")) Then
        End If
        If Not DA.QueryCommand(SQL_ULT_EST_PREC, VarkbufferEC0, lngHowManyEC0, MatCisterna) Then
                LogMessage "Error en el proceso de captura de precintos Thaumat. " & vbCrLf
                gridProcedures.RefetchCol 1
        End If
    End If
    If lngHowManyEC0 > 0 Then
    Latitud = Replace(Latitud, Chr(44), Chr(46))
    Longitud = Replace(Longitud, Chr(44), Chr(46))
    FechaEvento = FechaEv
    If Dep1 = "SE" Then Dep1 = "PR"
    If Dep2 = "SE" Then Dep2 = "PR"
    If Dep3 = "SE" Then Dep3 = "PR"
    If Dep4 = "SE" Then Dep4 = "PR"
    If Dep5 = "SE" Then Dep5 = "PR"
    If Dep6 = "SE" Then Dep6 = "PR"
    THDepo1 = Dep1
    THDepo2 = Dep2
    THDepo3 = Dep3
    THDepo4 = Dep4
    THDepo5 = Dep5
    THDepo6 = Dep5
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
                    "VALUES    ('" & MatCisterna & "', '" & FechaEvento & "', GETDATE(), 'T', '1', '1', '01', 'TR', '" & EDepCep & "', 'XX','" & Latitud & "','" & Longitud & "','1','" & Speed & "')"
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
                    "VALUES    ('" & MatCisterna & "', '" & FechaEvento & "', GETDATE(), 'T', '1', '1', '02', 'TR', '" & EDepCep & "', 'XX','" & Latitud & "','" & Longitud & "','1'," & Speed & ")"
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
                    "VALUES    ('" & MatCisterna & "', '" & FechaEvento & "', GETDATE(), 'T', '1', '1', '03', 'TR', '" & EDepCep & "', 'XX','" & Latitud & "','" & Longitud & "','1'," & Speed & ")"
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
                    "VALUES    ('" & MatCisterna & "', '" & FechaEvento & "', GETDATE(), 'T', '1', '1', '04', 'TR', '" & EDepCep & "', 'XX','" & Latitud & "','" & Longitud & "','1'," & Speed & ")"
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
                    "VALUES    ('" & MatCisterna & "', '" & FechaEvento & "', GETDATE(), 'T', '1', '1', '05', 'TR', '" & EDepCep & "', 'XX','" & Latitud & "','" & Longitud & "','1'," & Speed & ")"
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
                    "VALUES    ('" & MatCisterna & "', '" & FechaEvento & "', GETDATE(), 'T', '1', '1', '06', 'TR', '" & EDepCep & "', 'XX','" & Latitud & "','" & Longitud & "','1'," & Speed & ")"
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
        
        ' Borrar tabla position
'        If Not (DA.ExecuteQuery("DELETE FROM EVENTOS_THMAT WHERE (DATETIME < GETDATE() - 5)")) Then
'           LogMessage "Error borrado eventos: " & RegistroLeid & vbCrLf
'        End If
'
   
GoTo EndProc

ErrHandler:
    

    
EndProc:

    Captura_prec = 1

End Function
Public Function ImportaClickFind()


LogMessage "Inicio poceso importacion CLick & Find." & vbCrLf

XA.value(3, 1) = "En Ejecucion"
gridProcedures.RefetchCol 1

    Dim SQL_INSGIS As String
    Dim DA As CDataAccess
    Set DA = DataConnect.DA
    Dim SQL_Pos_error As String
    Dim ArchOrigen As String
    Dim linea, Dirtemp As String
    Dim contenido, car As String
    Dim directorio(0 To 1) As String
    Dim linfich, cont, codcar As Integer
    Dim Matricula, letra As String
    Dim x, j, pos As Integer
    Dim Valor(0 To 6) As String
    Dim FechaUTCVálida, FechaServ, FechaActSis, CordxStr, CordyStr As String
    Dim Cordx, Cordy  As Double
    Dim SQL_INSCLICK As String
    j = 0
    Dim ret As Boolean
    Dim xy As RESOIL2.CTextFile
    Dim SQL_NOLLEGA_CLICK As String
    Dim SQL_NOLLEGA_THMAT As String
    Dim rowbuff As Variant
    Dim lngnocap As Long
    Set xy = New CTextFile



    directorio(0) = DIR_CLICKFIND.Text + "ClickandFind\recibidos\"
    directorio(1) = DIR_CLICKFIND.Text + "ClickandFind\recibidos\historico\"

    Dirtemp = "temp\"



    ArchOrigen = Dir(directorio(0) + "*.txt")
            Do While ArchOrigen <> ""
                FileCopy (directorio(0) + ArchOrigen), (directorio(0) + Dirtemp + ArchOrigen)
                Kill (directorio(0) + ArchOrigen)
                ArchOrigen = Dir(directorio(0) + "*.txt")
            Loop

        ArchOrigen = Dir(directorio(0) + Dirtemp + "*.txt")

        Do While ArchOrigen <> ""
        

        
        'Abrimos el fichero, comprobamos que exista el final de carro y seguimos.
        Open (directorio(0) + Dirtemp + ArchOrigen) For Input As #1
        contenido = Input(LOF(1), #1)
        If Len(contenido) > 0 Then
        car = Mid(contenido, Len(contenido) - 1, 1)
        codcar = ASC(car)
        If codcar = 59 Then
        linfich = 0
            For cont = 1 To Len(contenido)
                linea = ""
                codcar = 0
                If cont > 1 Then cont = cont - 1
                Do While (codcar <> 10 And codcar <> 13) And cont <= Len(contenido)
                    car = Mid(contenido, cont, 1)
                    linea = linea + car
                    cont = cont + 1
                    codcar = ASC(car)
                Loop
            ' Aqui puedes coger los campos de la linea que necesites para guardarlos en un variant más manejable
      

                Valor(j) = ""
                    Do While pos < Len(linea) - 1
                        car = Mid(linea, pos + 1, 1)
                        pos = pos + 1
                        letra = car
                            
                        If (letra <> ";") Then
                            If (j < 7) Then
                                 Valor(j) = Valor(j) + car
                            End If
                           
                        Else
                                j = j + 1
                            If (j < 7) Then
                                Valor(j) = ""
                            End If
                        End If
                    Loop
                pos = 0
                Matricula = Valor(0)
                FechaUTCVálida = Format(Valor(1), "dd/mm/yyyy hh:mm:ss")
                
                CordxStr = Mid(Trim(Valor(3)), 1, 10)
                CordyStr = Mid(Trim(Valor(2)), 1, 10)
                CordxStr = Replace(CordxStr, ".", ",")
                CordyStr = Replace(CordyStr, ".", ",")
                Cordx = Round(CDbl(CordxStr), 5)
                Cordy = Round(CDbl(CordyStr), 5)
                Cordx = Left(Cordx, 8)
                Cordy = Left(Cordy, 8)
                j = 0
                    

'                MinsSinPOs = DateDiff("n", FechaUTCVálida, Now())
'                If MinsSinPOs < 2880 Then
'                ' Miramos la fecha de la última posicón válida que tenemos de la unidad
''                    SQL_Ult_Pos = "SELECT FEUTCVAL FROM PosicionesOesia WHERE (IDVEHIC = ?) AND (FEUTCVAL = ?)"
''                    If Not DA.QueryCommand(SQL_Ult_Pos, VarkbufferUltPOS, lngHowManyUltPOS, IdV, FechaUTCVálida) Then
''                    End If
''
''                    If lngHowManyUltPOS = 0 Then
''                        SQL_INSGIS = "Insert Into PosicionesOesia (IDVEHIC, FEUTCVAL, FEUTCULT, FESERV, FEACT, Longitud, Latitud) Values (" + IdV + ",'" + FechaUTCVálida + "','" + FechaUTCUltima + "','" + FechaServ + "','" + FechaUTCVálida + "','" + cordx + "','" + cordy + "')"
''                        If Not DA.ExecuteQuery(SQL_INSGIS) Then
''                             LogMessage "Error en la Importación de fichero: " + ArchOrigen + " El ID: " + IdV & vbCrLf
''                        End If
''                    End If
                    SQL_INSCLICK = "Insert Into PosicionesClick (MATRICULA, FECHAVAL, FECHAPOS, FECHALLE, LATITUD, LONGITUD) Values ('" + Matricula + "','" + FechaUTCVálida + "','" + FechaUTCVálida + "',getdate(),'" + CStr(Cordx) + "','" + CStr(Cordy) + "')"
                    If Not DA.ExecuteQuery(SQL_INSCLICK) Then
                         LogMessage "Error en la Importación de fichero: " + ArchOrigen + " El ID: " + Matricula & vbCrLf
                    End If
'
'                End If
            Next cont
         Close #1
         
          FileCopy (directorio(0) + Dirtemp + ArchOrigen), (directorio(1) + ArchOrigen)
          Kill (directorio(0) + Dirtemp + ArchOrigen)
        ArchOrigen = Dir(directorio(0) + Dirtemp + "*.txt")
       Else
       Close #1
       Kill (directorio(0) + Dirtemp + ArchOrigen)
        ArchOrigen = Dir(directorio(0) + Dirtemp + "*.txt")
       End If
       Else
        Close #1
        Kill (directorio(0) + Dirtemp + ArchOrigen)
        ArchOrigen = Dir(directorio(0) + Dirtemp + "*.txt")
       End If
       Loop
    
    ' Alarma no llegada posiciones Click
    SQL_NOLLEGA_CLICK = "SELECT TOP (1) FECHAVAL FROM PosicionesClick WHERE (FECHAVAL > GETDATE() - 0.08) AND (PROVEEDOR <> 'ST' OR PROVEEDOR IS NULL)"
    If DA.QueryCommand(SQL_NOLLEGA_CLICK, rowbuff, lngnocap) Then
        If (lngnocap = 0) Then
            ret = xy.OpenFile(MaskEdUnidad.Text & ":\Resoil2\Alarmas\POSICIONESCLICK.TXT", 1)
            ret = xy.CloseFile()
        End If
    End If

    ' Alarma no llegada posiciones Thmat
    SQL_NOLLEGA_THMAT = "SELECT TOP (1) DATETIME FROM EVENTOS_THMAT WHERE (DATETIME > GETDATE() - 0.8)"
    If DA.QueryCommand(SQL_NOLLEGA_THMAT, rowbuff, lngnocap) Then
        If (lngnocap = 0) Then
            ret = xy.OpenFile(MaskEdUnidad.Text & ":\Resoil2\Alarmas\POSICIONESTHMAT.TXT", 1)
            ret = xy.CloseFile()
        End If
    End If


LogMessage "Poceso de importacion CLick & FindFinalizado " & vbCrLf
XA.value(3, 1) = "En Espera"
gridProcedures.RefetchCol 1

End Function

Public Function ImportaStartit()


LogMessage "Inicio poceso importacion Startit." & vbCrLf

XA.value(3, 1) = "En Ejecucion"
gridProcedures.RefetchCol 1

    Dim SQL_INSGIS As String
    Dim DA As CDataAccess
    Set DA = DataConnect.DA
    Dim SQL_Pos_error As String
    Dim ArchOrigen As String
    Dim linea, Dirtemp As String
    Dim contenido, car As String
    Dim directorio(0 To 1) As String
    Dim linfich, cont, codcar As Integer
    Dim Matricula, letra As String
    Dim x, j, pos As Integer
    Dim Valor(0 To 10) As String
    Dim FechaUTCVálida, FechaServ, FechaActSis, CordxStr, CordyStr As String
    Dim Cordx, Cordy  As Double
    Dim SQL_INSCLICK As String
    j = 0
    Dim ret As Boolean
    Dim xy As RESOIL2.CTextFile
    Dim SQL_NOLLEGA_CLICK As String
    Dim SQL_NOLLEGA_THMAT As String
    Dim rowbuff As Variant
    Dim lngnocap As Long
    Set xy = New CTextFile
    Dim Proveedor As String
    Dim Velocidad As Integer
    Dim Dep1, Dep2, Dep3, Dep4, Dep5, Dep6 As String


    directorio(0) = DIR_CLICKFIND.Text + "Startit\recibidos\"
    directorio(1) = DIR_CLICKFIND.Text + "Startit\recibidos\historico\"

    Dirtemp = "temp\"



    ArchOrigen = Dir(directorio(0) + "*.*")
            Do While ArchOrigen <> ""
                FileCopy (directorio(0) + ArchOrigen), (directorio(0) + Dirtemp + ArchOrigen)
                Kill (directorio(0) + ArchOrigen)
                ArchOrigen = Dir(directorio(0) + "*.*")
            Loop

        ArchOrigen = Dir(directorio(0) + Dirtemp + "*.*")

        Do While ArchOrigen <> ""
        

        
        'Abrimos el fichero, comprobamos que exista el final de carro y seguimos.
        Open (directorio(0) + Dirtemp + ArchOrigen) For Input As #1
        contenido = Input(LOF(1), #1)
        If Len(contenido) > 0 Then
        car = Mid(contenido, Len(contenido) - 2, 1)
        codcar = ASC(car)
        If codcar = 59 Then
        linfich = 0
            For cont = 1 To Len(contenido)
                linea = ""
                codcar = 0
                If cont > 1 Then cont = cont - 1
                Do While (codcar <> 10) And cont <= Len(contenido)
                    car = Mid(contenido, cont, 1)
                    linea = linea + car
                    cont = cont + 1
                    codcar = ASC(car)
                Loop
            ' Aqui puedes coger los campos de la linea que necesites para guardarlos en un variant más manejable
      

                Valor(j) = ""
                    Do While pos < Len(linea) - 1
                        car = Mid(linea, pos + 1, 1)
                        pos = pos + 1
                        letra = car
                            
                        If (letra <> ";") Then
                            If (j < 11) Then
                                 Valor(j) = Valor(j) + car
                            End If
                           
                        Else
                                j = j + 1
                            If (j < 11) Then
                                Valor(j) = ""
                            End If
                        End If
                    Loop
                pos = 0
                Matricula = Valor(0)
                FechaUTCVálida = Format(Valor(1), "dd/mm/yyyy hh:mm:ss")
                
                CordxStr = Mid(Trim(Valor(3)), 1, 10)
                CordyStr = Mid(Trim(Valor(2)), 1, 10)
                CordxStr = Replace(CordxStr, ".", ",")
                CordyStr = Replace(CordyStr, ".", ",")
                Cordx = Round(CDbl(CordxStr), 5)
                Cordy = Round(CDbl(CordyStr), 5)
                Cordx = Left(Cordx, 8)
                Cordy = Left(Cordy, 8)
                j = 0
                    

'                MinsSinPOs = DateDiff("n", FechaUTCVálida, Now())
'                If MinsSinPOs < 2880 Then
'                ' Miramos la fecha de la última posicón válida que tenemos de la unidad
''                    SQL_Ult_Pos = "SELECT FEUTCVAL FROM PosicionesOesia WHERE (IDVEHIC = ?) AND (FEUTCVAL = ?)"
''                    If Not DA.QueryCommand(SQL_Ult_Pos, VarkbufferUltPOS, lngHowManyUltPOS, IdV, FechaUTCVálida) Then
''                    End If
''
''                    If lngHowManyUltPOS = 0 Then
''                        SQL_INSGIS = "Insert Into PosicionesOesia (IDVEHIC, FEUTCVAL, FEUTCULT, FESERV, FEACT, Longitud, Latitud) Values (" + IdV + ",'" + FechaUTCVálida + "','" + FechaUTCUltima + "','" + FechaServ + "','" + FechaUTCVálida + "','" + cordx + "','" + cordy + "')"
''                        If Not DA.ExecuteQuery(SQL_INSGIS) Then
''                             LogMessage "Error en la Importación de fichero: " + ArchOrigen + " El ID: " + IdV & vbCrLf
''                        End If
''                    End If
                    If Left(ArchOrigen, 4) = "3242" Then
                    Proveedor = "LC"
                    'llamar funcion rellenado tanques pasando el string valor
                        Velocidad = Valor(4)
                        Dep1 = Valor(5)
                        Dep2 = Valor(6)
                        Dep3 = Valor(7)
                        Dep4 = Valor(8)
                        Dep5 = Valor(9)
                        Dep6 = Valor(6)
                        If Captura_prec(Matricula, FechaUTCVálida, Cordy, Cordx, Velocidad, Dep1, Dep2, Dep3, Dep4, Dep5, Dep6) Then
                        End If
                    Else
                        If Left(ArchOrigen, 2) = "eq" Then
                        Proveedor = "EQ"
                        'Cambiar hora utc verano +2 invierno +1
                        FechaUTCVálida = DateAdd("h", 2, FechaUTCVálida)
                        FechaUTCVálida = Format(FechaUTCVálida, "dd/mm/yyyy hh:mm:ss")
                        
                        'llamar funcion rellenado tanques pasando el string valor
                        
                        Velocidad = Valor(4)
                        Dep1 = Valor(5)
                        Dep2 = Valor(6)
                        Dep3 = Valor(7)
                        Dep4 = Valor(8)
                        Dep5 = Valor(9)
                        Dep6 = Valor(6)
                        If Captura_prec(Matricula, FechaUTCVálida, Cordy, Cordx, Velocidad, Dep1, Dep2, Dep3, Dep4, Dep5, Dep6) Then
                        End If
                        Else
                        Proveedor = "ST"
                        End If
                    End If
                    
                    SQL_INSCLICK = "Insert Into PosicionesClick (MATRICULA, FECHAVAL, FECHAPOS, FECHALLE, LATITUD, LONGITUD, PROVEEDOR) Values ('" + Matricula + "','" + FechaUTCVálida + "','" + FechaUTCVálida + "',getdate(),'" + CStr(Cordx) + "','" + CStr(Cordy) + "','" + Proveedor + "')"
                    If Not DA.ExecuteQuery(SQL_INSCLICK) Then
                         LogMessage "Error en la Importación de fichero: " + ArchOrigen + " El ID: " + Matricula & vbCrLf
                    End If
'
'                End If
            Next cont
         Close #1
         
          FileCopy (directorio(0) + Dirtemp + ArchOrigen), (directorio(1) + ArchOrigen)
          Kill (directorio(0) + Dirtemp + ArchOrigen)
        ArchOrigen = Dir(directorio(0) + Dirtemp + "*.*")
       Else
       Close #1
       Kill (directorio(0) + Dirtemp + ArchOrigen)
        ArchOrigen = Dir(directorio(0) + Dirtemp + "*.*")
       End If
       Else
        Close #1
        Kill (directorio(0) + Dirtemp + ArchOrigen)
        ArchOrigen = Dir(directorio(0) + Dirtemp + "*.*")
       End If
       Loop
    
'    ' Alarma no llegada posiciones Click
'    SQL_NOLLEGA_CLICK = "SELECT TOP (1) FECHAVAL FROM PosicionesClick WHERE (FECHAVAL > GETDATE() - 0.08)"
'    If DA.QueryCommand(SQL_NOLLEGA_CLICK, rowbuff, lngnocap) Then
'        If (lngnocap = 0) Then
'            ret = xy.OpenFile(MaskEdUnidad.Text & ":\Resoil2\Alarmas\POSICIONESCLICK.TXT", 1)
'            ret = xy.CloseFile()
'        End If
'    End If
'
'    ' Alarma no llegada posiciones Thmat
'    SQL_NOLLEGA_THMAT = "SELECT TOP (1) DATETIME FROM EVENTOS_THMAT WHERE (DATETIME > GETDATE() - 0.8)"
'    If DA.QueryCommand(SQL_NOLLEGA_THMAT, rowbuff, lngnocap) Then
'        If (lngnocap = 0) Then
'            ret = xy.OpenFile(MaskEdUnidad.Text & ":\Resoil2\Alarmas\POSICIONESTHMAT.TXT", 1)
'            ret = xy.CloseFile()
'        End If
'    End If


LogMessage "Poceso de importacion Startit Finalizado " & vbCrLf
XA.value(3, 1) = "En Espera"
gridProcedures.RefetchCol 1

End Function


Public Function ExportaClickFind()
LogMessage "Inicio poceso exportacion CLick & Find." & vbCrLf

XA.value(4, 1) = "En Ejecucion"
gridProcedures.RefetchCol 1

'    --SM -JG Detección de ejecución de proccesos
Dim ret As Boolean
Dim xy As RESOIL2.CTextFile
Set xy = New CTextFile
ret = xy.OpenFile(MaskEdUnidad.Text & ":\Resoil2\Alarmas\EXCLICK.txt", 1)
ret = xy.CloseFile()



Dim DA As CDataAccess
Set DA = DataConnect.DA
Dim SQL_Click_Enviar As String
Dim varKBufferClick As Variant
Dim lngHowManyClick As Long
Dim Longitud, Latitud, Longitudgr, Latitudgr, FechaPos, RegistroLeid, MatCisternaTh, MatCisterna, Nomfi, Nomfihr As String
Dim Tanque1, Tanque2, Tanque3, Tanque4, Tanque5, Tanque6 As String
Dim Signolat, Signolong, i, LongitudCampo, Velocidad As Integer
Dim SQL_UPDATE_THMATCLICK As String
Nomfihr = Time()
Nomfihr = Format$(Nomfihr, "hhmmss")
Nomfi = Format$(Date, "DDMMYYYY") + Nomfihr + ".txt"
SQL_UPDATE_THMATCLICK = "UPDATE EVENTOS_THMAT SET ENVIADO = 1 WHERE (CODPOSITION = ?)"
SQL_Click_Enviar = "SELECT * FROM EVEN_THMAT_CLICK"

    If Not DA.QueryCommand(SQL_Click_Enviar, varKBufferClick, lngHowManyClick) Then
    End If
        If lngHowManyClick > 0 Then
        If mobjOutFileCLICK.OpenFile(DIR_CLICKFIND.Text + "ClickandFind\enviados\" + Nomfi, 1) Then
        End If
        If mobjOutFileSTARTIT.OpenFile(DIR_CLICKFIND.Text + "Startit\enviados\" + "ST" + Nomfi, 1) Then
        End If
    For i = 0 To lngHowManyClick - 1
        FechaPos = varKBufferClick(3, i)
        RegistroLeid = varKBufferClick(0, i)
        MatCisternaTh = varKBufferClick(1, i)
        If Not IsNull(varKBufferClick(2, i)) Then
            MatCisterna = varKBufferClick(2, i)
            Else
            MatCisterna = ""
            End If
        If Not (IsNull(varKBufferClick(5, i))) Then Latitudgr = varKBufferClick(5, i)
        If Not (IsNull(varKBufferClick(6, i))) Then Longitudgr = varKBufferClick(6, i)
        If varKBufferClick(7, i) = "S" Then Signolat = -1 Else Signolat = 1
        If varKBufferClick(8, i) = "W" Then Signolong = -1 Else Signolong = 1

        If Latitudgr <> "" And Longitudgr <> "" Then
            If Len(Longitudgr) < 9 Then
                LongitudCampo = Len(Longitudgr)
                Do While LongitudCampo < 9
                Longitudgr = "0" & Longitudgr
                LongitudCampo = LongitudCampo + 1
                Loop
            End If
            
            Latitud = CInt(Left(Latitudgr, 2)) + (CDbl((Mid(Latitudgr, 3, 2)) & "," & Right(Latitudgr, 4)) / 60)
            Longitud = CInt(Left(Longitudgr, 3)) + (CDbl((Mid(Longitudgr, 4, 2)) & "," & Right(Longitudgr, 4)) / 60)
            Latitud = Round(Latitud, 8)
            Longitud = Round(Longitud, 8)
            Latitud = Latitud * Signolat
            Longitud = Longitud * Signolong
            Latitud = Replace(Latitud, Chr(44), Chr(46))
            Longitud = Replace(Longitud, Chr(44), Chr(46))
        End If
        Velocidad = 0
        If Not (IsNull(varKBufferClick(9, i))) Then Velocidad = varKBufferClick(9, i)
        If Velocidad = 255 Then Velocidad = 0
        ' Convertimos a KM/H los nudos marinos
        Velocidad = Velocidad * 1.852
        Velocidad = Round(Velocidad)
        ' Envio CLick & Find
        
        mobjOutFileCLICK.Delimiter = ";"
        mobjOutFileCLICK.Item("1").value = MatCisterna
        mobjOutFileCLICK.Item("2").value = Format$(FechaPos, "yyyy-MM-dd") + " " + Format$(FechaPos, "hh:mm:ss")
        mobjOutFileCLICK.Item("3").value = Latitud
        mobjOutFileCLICK.Item("4").value = Longitud
        mobjOutFileCLICK.Item("5").value = Velocidad
        mobjOutFileCLICK.WriteLineNofij (True)
    
         ' Envio Startit
        
        If Not (IsNull(varKBufferClick(21, i))) Then Tanque1 = Right(RTrim(varKBufferClick(21, i)), 2)
        If Not (IsNull(varKBufferClick(22, i))) Then Tanque2 = Right(RTrim(varKBufferClick(22, i)), 2)
        If Not (IsNull(varKBufferClick(23, i))) Then Tanque3 = Right(RTrim(varKBufferClick(23, i)), 2)
        If Not (IsNull(varKBufferClick(24, i))) Then Tanque4 = Right(RTrim(varKBufferClick(24, i)), 2)
        If Not (IsNull(varKBufferClick(25, i))) Then Tanque5 = Right(RTrim(varKBufferClick(25, i)), 2)
        If Not (IsNull(varKBufferClick(26, i))) Then Tanque6 = Right(RTrim(varKBufferClick(26, i)), 2)
        
        mobjOutFileSTARTIT.Delimiter = ";"
        mobjOutFileSTARTIT.Item("1").value = MatCisterna
        mobjOutFileSTARTIT.Item("2").value = Format$(FechaPos, "yyyy-MM-dd") + " " + Format$(FechaPos, "hh:mm:ss")
        mobjOutFileSTARTIT.Item("3").value = Latitud
        mobjOutFileSTARTIT.Item("4").value = Longitud
        mobjOutFileSTARTIT.Item("5").value = Velocidad
        mobjOutFileSTARTIT.Item("6").value = Tanque1
        mobjOutFileSTARTIT.Item("7").value = Tanque2
        mobjOutFileSTARTIT.Item("8").value = Tanque3
        mobjOutFileSTARTIT.Item("9").value = Tanque4
        mobjOutFileSTARTIT.Item("10").value = Tanque5
        mobjOutFileSTARTIT.Item("11").value = Tanque6
        mobjOutFileSTARTIT.WriteLineNofij (True)
    
        If Not DA.ExecuteCommand(SQL_UPDATE_THMATCLICK, RegistroLeid) Then
        End If
    
    Next i
    mobjOutFileCLICK.CloseFile
    mobjOutFileSTARTIT.CloseFile
End If
LogMessage "Poceso de exportacion CLick & FindFinalizado " & vbCrLf
XA.value(4, 1) = "En Espera"
gridProcedures.RefetchCol 1
End Function

Public Function RellenaEsp(ByVal Cadena As String, ByVal LongCamp) As String

Dim longCadena As Integer
Dim Rellenar As Integer
Dim i As Integer

longCadena = Len(Cadena)

Rellenar = LongCamp - longCadena

If Rellenar > 0 Then
    i = 0
    
    Do While i < Rellenar
    
    Cadena = " " & Cadena
    
    i = i + 1
    Loop
    End If
    
RellenaEsp = Cadena

End Function

Public Function Rellenapos(Unidades As String)


Dim DA As CDataAccess
Set DA = DataConnect.DA
Dim h, j, k  As Integer
Dim VarkbufferPOS As Variant
Dim lngHowManyPOS As Long
Dim VarkbufferCPOS As Variant
Dim lngHowManyCPOS As Long
Dim VarkbufferUltPOS As Variant
Dim lngHowManyUltPOS As Long
Dim VarkbufferUltPOSVehi As Variant
Dim lngHowManyUltPOSVehi As Long
Dim SQL_BUSCA_TRACT As String
Dim Longitud, Latitud, Unidad, FechaPos, FechaPosVal, NuevafePOS, FechaUltPos As String
Dim Diferencia, MargenEspera, TiempoMargen, Posarellenar, SinFuncionar As Long
Dim SQL_CAPTURA_POS As String
Dim SQL_INSERTA_GIS As String
Dim SQL_Ultima_PosVal As String
Dim SQL_Ultima_PosValVEhi As String
Dim Rellanar As Boolean
Dim Rellenar As Boolean

Select Case Unidades

    Case "TECNO"
    
    SQL_BUSCA_TRACT = "SELECT TOP (1000) IDVEHIC, MAX(FEACT) AS FEACT, MAX(FEUTCVAL) AS FEUTCVAL FROM PosicionesOesia GROUP BY IDVEHIC HAVING  (IDVEHIC >= 0) AND (IDVEHIC < 800)"
    SQL_Ultima_PosVal = "SELECT MAX(FEUTCVAL) FROM PosicionesOesia WHERE (IDVEHIC > 0) and (IDVEHIC < 800)"
    SQL_CAPTURA_POS = "SELECT LATITUD, LONGITUD FROM OESIA.dbo.Posiciones WHERE (IDVEHIC = ?) AND (FEACT = ?)"

    Case "THMAT"
    
    SQL_BUSCA_TRACT = "SELECT TOP (1000) IDVEHIC, MAX(FEACT) AS FEACT, MAX(FEUTCVAL) AS FEUTCVAL FROM PosicionesOesia GROUP BY IDVEHIC HAVING  (IDVEHIC >= 800)"
    SQL_Ultima_PosVal = "SELECT MAX(FEUTCVAL) FROM PosicionesOesia WHERE (IDVEHIC >= 800)"
    SQL_CAPTURA_POS = "SELECT LATITUD, LONGITUD FROM OESIA.dbo.Posiciones WHERE (IDVEHIC = ?) AND (FEACT = ?)"
    
    Case "CLICK"
    
    SQL_BUSCA_TRACT = "SELECT     TOP (1000) MATRICULA, MAX(FECHAPOS) AS FECHAPOS , MAX(FECHAVAL) AS FECHAVAL FROM PosicionesClick GROUP BY MATRICULA HAVING (MATRICULA >= '0')"
    SQL_Ultima_PosVal = "SELECT MAX(FECHAVAL) AS FECHAVAL FROM PosicionesClick WHERE (FECHAVAL <= GETDATE())"
    SQL_Ultima_PosValVEhi = "SELECT MAX(FECHAVAL) AS FECHAVAL FROM PosicionesClick WHERE (MATRICULA = ?) AND(FECHAVAL <= GETDATE())"
    SQL_CAPTURA_POS = "SELECT  LATITUD, LONGITUD  FROM PosicionesClick WHERE (MATRICULA = ?) AND (FECHAPOS = ?)"

End Select

Rellenar = True
' Buscamos la última posición válida para detectar hasta cuando el sistema funciona

If Not DA.QueryCommand(SQL_Ultima_PosVal, VarkbufferUltPOS, lngHowManyUltPOS) Then
           End If
If lngHowManyUltPOS > 0 Then
    If Not IsNull(VarkbufferUltPOS(0, 0)) Then
    FechaUltPos = VarkbufferUltPOS(0, 0)
    Else
    FechaUltPos = "01/01/1900 00:00"
    End If
End If


If Not DA.QueryCommand(SQL_BUSCA_TRACT, VarkbufferPOS, lngHowManyPOS) Then
           End If

For h = 0 To lngHowManyPOS - 1

        Unidad = VarkbufferPOS(0, h)
        FechaPos = VarkbufferPOS(1, h)
        If Not DA.QueryCommand(SQL_Ultima_PosValVEhi, VarkbufferUltPOSVehi, lngHowManyUltPOSVehi, Unidad) Then
        End If
        If lngHowManyUltPOSVehi > 0 Then
        FechaPosVal = VarkbufferUltPOSVehi(0, 0)
        End If
        
        'FechaPosVal = FechaUltPos
        
        MargenEspera = MaskRepos.Text
        Diferencia = DateDiff("n", FechaPos, Now())
        SinFuncionar = DateDiff("n", FechaUltPos, Now())
        Posarellenar = Diferencia - MargenEspera
        
        
        'Solo rellenamos posiciones si la unidad lleva menos de 2 días sin funcionar
        'if sinfuncionar > 2880 then rellnar = false
        
        
        If Rellenar = True Then
            
                
            If Not DA.QueryCommand(SQL_CAPTURA_POS, VarkbufferCPOS, lngHowManyCPOS, Unidad, FechaPos) Then
            End If
            If lngHowManyCPOS > 0 Then
                Latitud = VarkbufferCPOS(0, 0)
                Longitud = VarkbufferCPOS(1, 0)
                NuevafePOS = FechaPos
                For k = 0 To Posarellenar
                NuevafePOS = DateAdd("n", 1, NuevafePOS)
              
               
                Select Case Unidades
                Case "TECNO"
                SQL_INSERTA_GIS = "Insert Into PosicionesOesia (IDVEHIC, FEUTCVAL, FEUTCULT, FESERV, FEACT, Latitud, Longitud) Values (" & Unidad & ",'" & FechaPosVal & "','" & FechaPosVal & "',getdate(),'" & NuevafePOS & "','" & Latitud & "','" & Longitud & "')"
                Case "THMAT"
                SQL_INSERTA_GIS = "Insert Into PosicionesOesia (IDVEHIC, FEUTCVAL, FEUTCULT, FESERV, FEACT, Latitud, Longitud) Values (" & Unidad & ",'" & FechaPosVal & "','" & FechaPosVal & "',getdate(),'" & NuevafePOS & "','" & Latitud & "','" & Longitud & "')"
                Case "CLICK"
                SQL_INSERTA_GIS = "Insert Into Posicionesclick (MATRICULA, FECHAVAL, FECHALLE, FECHAPOS, Latitud, Longitud, Proveedor) Values ('" & Unidad & "','" & FechaPosVal & "',getdate(),'" & NuevafePOS & "','" & Latitud & "','" & Longitud & "','1')"
                End Select
                If Not DA.ExecuteQuery(SQL_INSERTA_GIS) Then
                End If
                 Next k
            End If
        End If
        
Next h


End Function
Public Function RellenaposVac(Unidades As String)


Dim DA As CDataAccess
Set DA = DataConnect.DA
Dim h, j, k  As Integer
Dim VarkbufferPOS As Variant
Dim lngHowManyPOS As Long
Dim VarkbufferCPOS As Variant
Dim lngHowManyCPOS As Long
Dim VarkbufferUltPOS As Variant
Dim lngHowManyUltPOS As Long
Dim SQL_Pos_Rellenar As String
Dim Longitud, Latitud, Unidad, UnidadAnt, FechaPos, FechaPosVal, NuevafePOS, FechaUltPos As String
Dim Diferencia, MargenEspera, TiempoMargen, Posarellenar, SinFuncionar As Long
Dim SQL_CAPTURA_POS As String
Dim SQL_INSERTA_GIS As String
Dim SQL_Ultima_PosVal As String
Dim SQL_UltimaPOS_Vehi As String

Select Case Unidades

   
    Case "CLICK"
    
    SQL_Pos_Rellenar = "SELECT  MATRICULA, horar FROM PosicionesClick_Arellenar ORDER BY MATRICULA DESC, horar"
    SQL_Ultima_PosVal = "SELECT MAX(FECHAVAL) AS FECHAVAL FROM PosicionesClick WHERE (FECHAVAL < GETDATE())"
    SQL_CAPTURA_POS = "SELECT  LATITUD, LONGITUD  FROM PosicionesClick WHERE (MATRICULA = ?) AND (FECHAPOS = ?)"
    SQL_UltimaPOS_Vehi = "SELECT  TOP (1) FECHAVAL, LATITUD, LONGITUD FROM PosicionesClick WHERE   (MATRICULA = ?) AND (FECHAPOS < ?) ORDER BY FECHAPOS DESC"

End Select

' Buscamos la última posición válida para detectar hasta cuando el sistema funciona

If Not DA.QueryCommand(SQL_Ultima_PosVal, VarkbufferUltPOS, lngHowManyUltPOS) Then
           End If
If lngHowManyUltPOS > 0 Then
    If Not IsNull(VarkbufferUltPOS(0, 0)) Then
    FechaUltPos = VarkbufferUltPOS(0, 0)
    Else
    FechaUltPos = "01/01/1900 00:00"
    End If
End If

SinFuncionar = DateDiff("n", FechaUltPos, Now())
' Si hace más de 120 minutos que no llegan posiciones validas no rellenamos
If SinFuncionar < 1111120 Then

    If Not DA.QueryCommand(SQL_Pos_Rellenar, VarkbufferPOS, lngHowManyPOS) Then
               End If
    
    For h = 0 To lngHowManyPOS - 1
    
            Unidad = VarkbufferPOS(0, h)
            UnidadAnt = Unidad
            FechaPos = Format(Now, "d/M/yyyy") & " " & VarkbufferPOS(1, h)
                If Not DA.QueryCommand(SQL_UltimaPOS_Vehi, VarkbufferCPOS, lngHowManyCPOS, Unidad, FechaPos) Then
                End If
                If lngHowManyCPOS > 0 Then
                FechaPosVal = VarkbufferCPOS(0, 0)
                Latitud = VarkbufferCPOS(1, 0)
                Longitud = VarkbufferCPOS(2, 0)
                End If
            Do While Unidad = UnidadAnt And h < lngHowManyPOS
                ' Controlamos el salto de fecha de las 00:00 para ello en el query debemos tener fecha -0.09375 (dos horas y 15 mins exactas) COmprobado funciona desde la 21:43
                If VarkbufferPOS(1, h) > "21:42" Then
                NuevafePOS = Format(Now - 1, "d/M/yyyy") & " " & VarkbufferPOS(1, h)
                Else
                NuevafePOS = Format(Now, "d/M/yyyy") & " " & VarkbufferPOS(1, h)
                End If
            SQL_INSERTA_GIS = "Insert Into Posicionesclick (MATRICULA, FECHAVAL, FECHALLE, FECHAPOS, Latitud, Longitud, Proveedor) Values ('" & Unidad & "','" & FechaPosVal & "',getdate(),'" & NuevafePOS & "','" & Latitud & "','" & Longitud & "','2')"
            If Not DA.ExecuteQuery(SQL_INSERTA_GIS) Then
            End If

            
            h = h + 1
            If h < lngHowManyPOS Then Unidad = VarkbufferPOS(0, h)
            If Unidad <> UnidadAnt Then h = h - 1
            Loop
'            'Solo rellenamos posiciones si la unidad lleva menos de 2 días sin funcionar
'
'                If Not DA.QueryCommand(SQL_CAPTURA_POS, VarkbufferCPOS, lngHowManyCPOS, Unidad, FechaPos) Then
'                End If
'                If lngHowManyCPOS > 0 Then
'                    Latitud = VarkbufferCPOS(0, 0)
'                    Longitud = VarkbufferCPOS(1, 0)
'                    NuevafePOS = FechaPos
'                    For k = 0 To Posarellenar
'                    NuevafePOS = DateAdd("n", 1, NuevafePOS)
'                    Select Case Unidades
'                    Case "TECNO"
'                    SQL_INSERTA_GIS = "Insert Into PosicionesOesia (IDVEHIC, FEUTCVAL, FEUTCULT, FESERV, FEACT, Latitud, Longitud) Values (" & Unidad & ",'" & FechaPosVal & "','" & FechaPosVal & "',getdate(),'" & NuevafePOS & "','" & Latitud & "','" & Longitud & "')"
'                    Case "THMAT"
'                    SQL_INSERTA_GIS = "Insert Into PosicionesOesia (IDVEHIC, FEUTCVAL, FEUTCULT, FESERV, FEACT, Latitud, Longitud) Values (" & Unidad & ",'" & FechaPosVal & "','" & FechaPosVal & "',getdate(),'" & NuevafePOS & "','" & Latitud & "','" & Longitud & "')"
'                    Case "CLICK"
'                    SQL_INSERTA_GIS = "Insert Into Posicionesclick (MATRICULA, FECHAVAL, FECHALLE, FECHAPOS, Latitud, Longitud) Values ('" & Unidad & "','" & FechaPosVal & "',getdate(),'" & NuevafePOS & "','" & Latitud & "','" & Longitud & "')"
'                    End Select
'                    If Not DA.ExecuteQuery(SQL_INSERTA_GIS) Then
'                    End If
'                     Next k
'                End If
'            End If
            
    Next h
End If

End Function

Private Sub Chkdesp_Click()

End Sub


Private Function SatBorraPLanif(ByVal Datos As Integer)

Dim DA As CDataAccess
Set DA = DataConnect.DA

Dim SQLBorradoSAt As String
Dim SQLBorradoVehi As String
Dim SQLAborrar As String
Dim varKBufferAborrar As Variant
Dim lngHowManyAborrar As Long
Dim i As Integer
Dim DepBorrar As String
Dim FechaBorrar As String

SQLAborrar = "SELECT        ORDIOCT0.OCTIDADE, ORDIOCT0.OCTDTCRI " & _
"FROM            ORDIOCT0 INNER JOIN " & _
"                         ORDIOCR0 ON ORDIOCT0.OCTIDOCT = ORDIOCR0.OCRIDOCT LEFT OUTER JOIN " & _
"                         SATURACION ON ORDIOCT0.OCTIDOCT = SATURACION.SAROCT " & _
"WHERE        (ORDIOCT0.OCTIDASC = '038') AND (SATURACION.SAROCT IS NULL) AND (ORDIOCT0.OCTFLSCA = 'N') AND (ORDIOCR0.OCRQRILT < 9500) " & _
"GROUP BY ORDIOCT0.OCTIDADE, ORDIOCT0.OCTDTCRI " & _
"HAVING        (ORDIOCT0.OCTIDADE = '004' OR " & _
"                         ORDIOCT0.OCTIDADE = '005' OR " & _
"                         ORDIOCT0.OCTIDADE = '008' OR " & _
"                         ORDIOCT0.OCTIDADE = '010' OR " & _
"                         ORDIOCT0.OCTIDADE = '002' OR " & _
"                         ORDIOCT0.OCTIDADE = '028' OR " & _
"                         ORDIOCT0.OCTIDADE = '016' OR " & _
"                         ORDIOCT0.OCTIDADE = '021' OR " & _
"                         ORDIOCT0.OCTIDADE = '030' OR " & _
"                         ORDIOCT0.OCTIDADE = '031' OR " & _
"                         ORDIOCT0.OCTIDADE = '032' OR " & _
"                         ORDIOCT0.OCTIDADE = '033' OR " & _
"                         ORDIOCT0.OCTIDADE = '034' OR " & _
"                         ORDIOCT0.OCTIDADE = '035') AND (ORDIOCT0.OCTDTCRI > GETDATE() -1) " & _
"ORDER BY ORDIOCT0.OCTDTCRI, ORDIOCT0.OCTIDADE "




Select Case Datos

    Case 0
    SQLBorradoSAt = "DELETE FROM SATURACION"
    SQLBorradoVehi = "DELETE FROM SATURA_VEHICULO"
    
    If Not DA.ExecuteQuery(SQLBorradoSAt) Then
    End If
    If Not DA.ExecuteQuery(SQLBorradoVehi) Then
    End If
   
    Case 1
    
    If DA.QueryCommand(SQLAborrar, varKBufferAborrar, lngHowManyAborrar) Then
    End If

    If lngHowManyAborrar > 0 Then
    i = 0
    Do While i < lngHowManyAborrar
    DepBorrar = varKBufferAborrar(0, i)
    FechaBorrar = varKBufferAborrar(1, i)
    FechaBorrar = Format(FechaBorrar, "yyyy-mm-dd hh:mm:ss")
    'Borra solo el depósito que tiene nuevas ordenes
    'SQLBorradoSAt = "DELETE FROM SATURACION WHERE (SARFECHA = CONVERT(DATETIME,'" & FechaBorrar & "',102)) AND (SARADE = '" & DepBorrar & "')"
    'SQLBorradoVehi = "DELETE FROM SATURA_VEHICULO  WHERE (SAVFECHA = CONVERT(DATETIME,'" & FechaBorrar & "',102)) AND (SAVDADE = '" & DepBorrar & "')"
    
    
    ' Borra todo
    SQLBorradoSAt = "DELETE FROM SATURACION WHERE (SARFECHA = CONVERT(DATETIME,'" & FechaBorrar & "',102))"
    SQLBorradoVehi = "DELETE FROM SATURA_VEHICULO  WHERE (SAVFECHA = CONVERT(DATETIME,'" & FechaBorrar & "',102))"
       
    If Not DA.ExecuteQuery(SQLBorradoSAt) Then
    End If
    If Not DA.ExecuteQuery(SQLBorradoVehi) Then
    End If
       
    i = i + 1
    Loop
    End If



End Select


End Function

Public Function SatBuscaVehiculo(ByVal FechaOrden As String, ByVal Depoorden As String, ByVal NOrden As String, ByVal CantOrden As Long, ByRef Vehiculoasign As String, ByRef Nviaje As Integer, ByRef Ndescarga As Integer, ByRef Motivo As String, ByRef Longx As String, ByRef Laty As String, ByRef Horaini As Date, ByRef HoraLlegada As Date, ByVal Bomba As Boolean, ByVal TanqueDedicado As String)

Dim DA As CDataAccess
Set DA = DataConnect.DA

Dim SQLVehiculos As String
Dim varKBufferVehiculos As Variant
Dim lngHowManyVehiculos As Long
Dim V As Integer
Dim Cantdisp As Long
Dim CantVehi As Long
Dim cantserv As Long
Dim Encontrado As Boolean
Dim Vehiculoleido As String
Dim Tipovehi As String
Dim Bombavehi As Boolean
Dim Cabe, CumpleBomba As Boolean
Dim TiempoTrabdia As Integer
Dim Excesotiempo As Boolean
Dim TiempoMAxtrab As Integer
'Tiempo maximo que trabaja un vehiculo
'TiempoMAxtrab = 720

SQLVehiculos = "SELECT        ANAGAUT0.AUTIDAUT, ANAGAUT0.AUTIDAM1, ANAGAUT0.AUTIDAM2, ANAGAUT0.AUTIDAU1, ANAGAUT0.AUTIDADE, ANAGAAM0.AAMIDATM, " & _
"                         CASE WHEN ANAGAAM0.AAMCAP01 > 0 THEN ANAGAAM0.AAMCAP01 ELSE ANAGAAM0_1.AAMCAP01 END AS CAP1, " & _
"                         CASE WHEN ANAGAAM0.AAMCAP02 > 0 THEN ANAGAAM0.AAMCAP02 ELSE ANAGAAM0_1.AAMCAP02 END AS CAP2, " & _
"                        CASE WHEN ANAGAAM0.AAMCAP03 > 0 THEN ANAGAAM0.AAMCAP03 ELSE ANAGAAM0_1.AAMCAP03 END AS CAP3, " & _
"                         CASE WHEN ANAGAAM0.AAMCAP04 > 0 THEN ANAGAAM0.AAMCAP04 ELSE ANAGAAM0_1.AAMCAP04 END AS CAP4, " & _
"                         CASE WHEN ANAGAAM0.AAMCAP05 > 0 THEN ANAGAAM0.AAMCAP05 ELSE ANAGAAM0_1.AAMCAP05 END AS CAP5, " & _
"                        CASE WHEN ANAGAAM0.AAMCAP06 > 0 THEN ANAGAAM0.AAMCAP06 ELSE ANAGAAM0_1.AAMCAP06 END AS CAP6, " & _
"                         CASE WHEN (ANAGAAM0.AAMCAP01 + ANAGAAM0.AAMCAP02 + ANAGAAM0.AAMCAP03 + ANAGAAM0.AAMCAP04 + ANAGAAM0.AAMCAP05 + ANAGAAM0.AAMCAP06) " & _
"                         > 0 THEN (ANAGAAM0.AAMCAP01 + ANAGAAM0.AAMCAP02 + ANAGAAM0.AAMCAP03 + ANAGAAM0.AAMCAP04 + ANAGAAM0.AAMCAP05 + ANAGAAM0.AAMCAP06) " & _
"                         ELSE (ANAGAAM0_1.AAMCAP01 + ANAGAAM0_1.AAMCAP02 + ANAGAAM0_1.AAMCAP03 + ANAGAAM0_1.AAMCAP04 + ANAGAAM0_1.AAMCAP05 + ANAGAAM0_1.AAMCAP06) END AS CAPACIDAD, " & _
"                             (SELECT        SUM(ORDIOCR0.OCRQRILT) AS Litros FROM SATURACION INNER JOIN ORDIOCT0 ON SATURACION.SAROCT = ORDIOCT0.OCTIDOCT INNER JOIN " & _
"                         ORDIOCR0 ON ORDIOCT0.OCTIDOCT = ORDIOCR0.OCRIDOCT WHERE (SATURACION.SARUT = SATURADISP.SATUT) AND (SATURACION.SARFECHA = SATURADISP.SATFECHA) " & _
" GROUP BY SATURACION.SARFECHA, SATURACION.SARUT) AS LtrTR,(SELECT        MAX(SARNVI) AS NVI FROM            SATURACION AS SATURACION_2 " & _
"                               GROUP BY SARFECHA, SARUT " & _
"                               HAVING         (SARUT = SATURADISP.SATUT) AND (SARFECHA = SATURADISP.SATFECHA)) AS Ultvi, " & _
"                             (SELECT        MAX(SARNDESC) AS NVI " & _
"                               FROM            SATURACION AS SATURACION_2 " & _
"                               GROUP BY SARFECHA, SARUT " & _
"                               HAVING         (SARUT = SATURADISP.SATUT) AND (SARFECHA = SATURADISP.SATFECHA)) AS Ultdesc, SATURADISP.SATHINI, " & _
"                               CASE WHEN ANAGAAM0.AAMIDATS <> '' THEN ANAGAAM0.AAMIDATS ELSE ANAGAAM0_1.AAMIDATS END AS Tipovehi,SATURADISP.SATUSO " & _
"FROM            SATURADISP LEFT OUTER JOIN " & _
"                         ANAGAUT0 ON SATURADISP.SATUT = ANAGAUT0.AUTIDAUT LEFT OUTER JOIN " & _
"                         ANAGAAM0 ON ANAGAUT0.AUTIDAM1 = ANAGAAM0.AAMIDAAM LEFT OUTER JOIN " & _
"                         ANAGAAM0 AS ANAGAAM0_1 ON ANAGAUT0.AUTIDAM2 = ANAGAAM0_1.AAMIDAAM " & _
"WHERE        (SATURADISP.SATFECHA = CONVERT(DATETIME,'" & FechaOrden & "' , 102)) and ANAGAUT0.AUTIDADE = '" & Depoorden & "'"

If DA.QueryCommand(SQLVehiculos, varKBufferVehiculos, lngHowManyVehiculos) Then
End If


Motivo = ""
V = 0
Encontrado = False
Cabe = False
CumpleBomba = False
If lngHowManyVehiculos > 0 Then

    Horaini = varKBufferVehiculos(16, V)
    Do While V < lngHowManyVehiculos And Encontrado = False
        TiempoMAxtrab = varKBufferVehiculos(18, V) * 60
        Excesotiempo = False
        Cabe = False
        CantVehi = varKBufferVehiculos(12, V)
        If IsNull(varKBufferVehiculos(14, V)) Then
            Nviaje = 0
        Else
            Nviaje = varKBufferVehiculos(14, V)
        End If
        Vehiculoleido = varKBufferVehiculos(0, V)
        
        If Not IsNull(varKBufferVehiculos(17, V)) Then
            Tipovehi = varKBufferVehiculos(17, V)
            Else
            Tipovehi = ""
        End If
    
        If Tipovehi = "002" Or Tipovehi = "006" Then
        Bombavehi = True
        Else
        Bombavehi = False
        End If
        
        If SatDatosViVehi(Vehiculoleido, FechaOrden, Nviaje, Ndescarga, cantserv, Longx, Laty, HoraLlegada, TiempoTrabdia) Then
        End If

    
        If TiempoTrabdia >= TiempoMAxtrab Then Excesotiempo = True
        Cantdisp = CantVehi - cantserv
        Dim Viajebuscar As Integer
        If Nviaje = 0 Then
        Viajebuscar = 1
        Else
        Viajebuscar = Nviaje
        End If
        If SatCHKTancaje(FechaOrden, NOrden, Vehiculoleido, Viajebuscar, TanqueDedicado, False) = True And Excesotiempo = False Then Cabe = True
        If Bombavehi = True And Bomba = True Then
            CumpleBomba = True
        Else
            If Bomba = False Then CumpleBomba = True
        End If
        
        ' Cumple bomba
        If Cabe = True And CumpleBomba = True Then Encontrado = True
        ' Cabe
        'If Cabe = True Then Encontrado = True
        V = V + 1
    
    Loop
    
    
    ' Si encontrado lo añadimos
    
    If Encontrado = True Then
        V = V - 1
        Vehiculoasign = varKBufferVehiculos(0, V)
        If Nviaje = 0 Then Nviaje = 1
        If SatUltimaDescVehi(Vehiculoasign, FechaOrden, Nviaje, Ndescarga) Then
        End If
        Ndescarga = Ndescarga + 1

     ' Si no encontrado buscamos otro vehículo para añadi otro viaje
    Else
        If CumpleBomba = True Then

            V = 0
            Do While V < lngHowManyVehiculos And Encontrado = False
                TiempoMAxtrab = varKBufferVehiculos(18, V) * 60
                Excesotiempo = False
                Encontrado = False
                CantVehi = varKBufferVehiculos(12, V)
                If IsNull(varKBufferVehiculos(14, V)) Then
                    Nviaje = 0
                Else
                    Nviaje = varKBufferVehiculos(14, V)
                End If
                Vehiculoleido = varKBufferVehiculos(0, V)
                
                If Not IsNull(varKBufferVehiculos(17, V)) Then
                    Tipovehi = varKBufferVehiculos(17, V)
                    Else
                    Tipovehi = ""
                End If
            
                If Tipovehi = "002" Or Tipovehi = "006" Then
                Bombavehi = True
                Else
                Bombavehi = False
                End If
                
                If SatDatosViVehi(Vehiculoleido, FechaOrden, Nviaje, Ndescarga, cantserv, Longx, Laty, HoraLlegada, TiempoTrabdia) Then
                End If
            
                Cantdisp = CantVehi - cantserv
        
                 If Bombavehi = True And Bomba = True Then
                 CumpleBomba = True
                 Else
                 If Bomba = False Then CumpleBomba = True
                 End If

                    
                ' Cumple bomba
                If TiempoTrabdia >= TiempoMAxtrab Then Excesotiempo = True
                Viajebuscar = Nviaje + 1
                If SatCHKTancaje(FechaOrden, NOrden, Vehiculoleido, Viajebuscar, TanqueDedicado, False) = True And CumpleBomba = True And Excesotiempo = False Then
                    Encontrado = True
                    Vehiculoasign = varKBufferVehiculos(0, V)
                End If
        
        
                V = V + 1
            Loop
                    
'            Do While V < lngHowManyVehiculos And Encontrado = False
'            Vehiculoasign = varKBufferVehiculos(0, 0)
'            If IsNull(varKBufferVehiculos(14, 0)) Then
'                Nviaje = 0
'            Else
'                Nviaje = varKBufferVehiculos(14, 0)
'            End If
'            If SatDatosViVehi(Vehiculoasign, FechaOrden, Nviaje, Ndescarga, cantserv, Longx, Laty, HoraLlegada) Then
'            End If
'            Loop
            If Excesotiempo = False And Encontrado = True Then
            Ndescarga = Ndescarga + 1
            Nviaje = Nviaje + 1
            Ndescarga = 1
            Else
                If Excesotiempo = True Then
                Motivo = "Los vehiculos del dep. " & Depoorden & " exceden el tiempo de trabajo"
                Else
                Motivo = "La orden no cabe por tancaje en ningun vehiculo del deposito  " & Depoorden
                End If
            Vehiculoasign = "ZZZ"
            End If
            
        Else
        Motivo = "Ningún vehículo con bomba en dep:" & Depoorden
        Vehiculoasign = "ZZZ"
        End If
        
    End If
Else
    Motivo = "No hay vehículos asignados en dep:" & Depoorden
    Vehiculoasign = "ZZZ"
End If

   
End Function

Public Function SAtCalcDistCoordenadas(ByVal CoorxLongDesde As Double, ByVal CoordYLatDesde As Double, ByVal CoorxLongHasta As Double, ByVal CoordYLatHasta As Double) As Double

Dim Distancia, Lat1, Long1, Lat2, Long2 As Double
Dim Radio As Double
Dim Multradio As Double
Dim PI As Double
Dim Longitud As Double
Dim Var As String
Dim Latitud

PI = 3.14159265358979

Radio = 6378.137
Multradio = PI / 180

Lat1 = CoordYLatDesde
Long1 = CoorxLongDesde
Lat2 = CoordYLatHasta
Long2 = CoorxLongHasta


Lat1 = Lat1 * Multradio
Long1 = Long1 * Multradio
Lat2 = Lat2 * Multradio
Long2 = Long2 * Multradio
Longitud = Long2 - Long1
Latitud = Lat2 - Lat1
If Longitud <> 0 And Latitud <> 0 Then
Var = Sin(Lat1) * Sin(Lat2) + Cos(Lat1) * Cos(Lat2) * Cos(Longitud)
Distancia = (ATN(-Var / Sqr(-Var * Var + 1)) + 2 * ATN(1)) * Radio
Else
Distancia = 0
End If


SAtCalcDistCoordenadas = Distancia

End Function

Private Function SatCHKTancaje(ByVal FechaOrden As String, ByVal Orden As String, ByVal Vehiculo As String, ByVal Nviaje As Integer, ByVal TanqueDedicado As String, ByVal Guardar As Boolean) As Boolean
Dim DA As CDataAccess
Set DA = DataConnect.DA


Dim SQLTancajes As String
Dim varKBufferTancajes As Variant
Dim lngHowManyTancajes As Long
Dim SQLInsertTanc As String
Dim SQLProdordenes As String
Dim VarKbufferProdordenes  As Variant
Dim LngHowmanyProdordenes  As Long
Dim SQLTancvehi As String
Dim VarKbufferTancvehi  As Variant
Dim LngHowmanyTancvehi  As Long
Dim P, T As Integer
Dim Cabe, CabeProd As Boolean
Dim LitrosOrd As Long
Dim LitrosRest As Long
Dim ProdOrd As String
Dim CantTanc As Long
Dim CantTrans As Long
Dim ProdTrans As String
Dim CantTransfin As Long
Dim SqlUpdateTanques As String
Dim Ntanque As Integer
Dim NtanqueRep As Integer

SQLTancajes = "SELECT        SAVFECHA, SAVUT, SAVNVI, SAVDEP1, SAVPROD1, SAVDEP2, SAVPROD2, SAVDEP3, SAVPROD3, SAVDEP4, SAVPROD4, SAVDEP5, SAVPROD5, SAVDEP6, SAVPROD6 " & _
"FROM            SATURA_VEHICULO " & _
"WHERE        (SAVFECHA = CONVERT(DATETIME, '" & FechaOrden & "', 102)) AND (SAVUT = '" & Vehiculo & "') AND (SAVNVI = '" & Nviaje & "')"




If lngHowManyTancajes < 1 Then

    ' No hay nada cargado por tanto añadimos la linea de momento solo miramos si cabe por litros
    
        
        SQLInsertTanc = " INSERT INTO SATURA_VEHICULO " & _
    "                         (SAVFECHA, SAVNVI, SAVUT, SAVDEP1CAP, SAVDEP2CAP, SAVDEP3CAP, SAVDEP4CAP, SAVDEP5CAP, SAVDEP6CAP, SAVDADE) " & _
    "SELECT        CONVERT(DATETIME,'" & FechaOrden & "',102) AS Fecha, " & Nviaje & " AS VI, ANAGAUT0.AUTIDAUT, CASE WHEN ANAGAAM0.AAMCAP01 > 0 THEN ANAGAAM0.AAMCAP01 ELSE ANAGAAM0_1.AAMCAP01 END AS CAP1, " & _
    "                        CASE WHEN ANAGAAM0.AAMCAP02 > 0 THEN ANAGAAM0.AAMCAP02 ELSE ANAGAAM0_1.AAMCAP02 END AS CAP2, " & _
    "                        CASE WHEN ANAGAAM0.AAMCAP03 > 0 THEN ANAGAAM0.AAMCAP03 ELSE ANAGAAM0_1.AAMCAP03 END AS CAP3, " & _
    "                         CASE WHEN ANAGAAM0.AAMCAP04 > 0 THEN ANAGAAM0.AAMCAP04 ELSE ANAGAAM0_1.AAMCAP04 END AS CAP4, " & _
    "                        CASE WHEN ANAGAAM0.AAMCAP05 > 0 THEN ANAGAAM0.AAMCAP05 ELSE ANAGAAM0_1.AAMCAP05 END AS CAP5, " & _
    "                        CASE WHEN ANAGAAM0.AAMCAP06 > 0 THEN ANAGAAM0.AAMCAP06 ELSE ANAGAAM0_1.AAMCAP06 END AS CAP6, ANAGAUT0.AUTIDADE " & _
    "FROM            ANAGAUT0 LEFT OUTER JOIN" & _
    "                         ANAGAAM0 ON ANAGAUT0.AUTIDAM1 = ANAGAAM0.AAMIDAAM LEFT OUTER JOIN " & _
    "                         ANAGAAM0 AS ANAGAAM0_1 ON ANAGAUT0.AUTIDAM2 = ANAGAAM0_1.AAMIDAAM " & _
    "WHERE        (ANAGAUT0.AUTIDAUT = '" & Vehiculo & "') "
        
        If Not DA.ExecuteQuery(SQLInsertTanc) Then
        End If
        

End If
    
    ' cargamos las lineas del pedido, dependiendo si se pueden juntar agrupamos o no
    
    If TanqueDedicado = "N" Then
    SQLProdordenes = "SELECT        OCRIDAPR, SUM(OCRQRILT) AS Expr1 " & _
                        "FROM            ORDIOCR0 " & _
                        "WHERE        (OCRIDOCT = '" & Orden & "') " & _
                        "GROUP BY OCRIDAPR "
    Else
    SQLProdordenes = "SELECT        OCRIDAPR, OCRQRILT " & _
                    "FROM            ORDIOCR0 " & _
                    "WHERE        (OCRIDOCT = '" & Orden & "')"
    
    End If

    If DA.QueryCommand(SQLProdordenes, VarKbufferProdordenes, LngHowmanyProdordenes) Then
    End If
    
    If LngHowmanyProdordenes > 0 Then

        ' Buscamos si cabe en el vehículo y viaje
    
        SQLTancvehi = "SELECT        SAVDEP1CAP, SAVDEP1, SAVPROD1, SAVDEP2CAP, SAVDEP2, SAVPROD2, SAVDEP3CAP, SAVDEP3, SAVPROD3, SAVDEP4CAP, SAVDEP4, SAVPROD4, SAVDEP5CAP, SAVDEP5, SAVPROD5, SAVDEP6CAP, " & _
                        "                         SAVDEP6, SAVPROD6 " & _
                        "FROM            SATURA_VEHICULO " & _
                        "WHERE        (SAVFECHA = CONVERT(DATETIME,'" & FechaOrden & "',102)) AND (SAVUT = '" & Vehiculo & "') AND (SAVNVI = " & Nviaje & ")"
    
        If DA.QueryCommand(SQLTancvehi, VarKbufferTancvehi, LngHowmanyTancvehi) Then
        End If
    
        If LngHowmanyTancvehi > 0 Then
            ' Recorremos para cada orden los tanques hasta que encontremos el que quema o lleguemos al final
            P = 0
            
            
            
            CabeProd = True
            Do While P < LngHowmanyProdordenes And CabeProd = True
            ProdOrd = VarKbufferProdordenes(0, P)
            LitrosOrd = VarKbufferProdordenes(1, P)
                'Recorremos los tanques
                T = 0
                Ntanque = 1
                Cabe = False
                Do While T < 17 And Cabe = False
                If Not IsNull(VarKbufferTancvehi(T, 0)) Then CantTanc = VarKbufferTancvehi(T, 0) Else CantTanc = 0
                If Not IsNull(VarKbufferTancvehi(T + 1, 0)) Then CantTrans = VarKbufferTancvehi(T + 1, 0) Else CantTrans = 0
                If Not IsNull(VarKbufferTancvehi(T + 2, 0)) Then ProdTrans = VarKbufferTancvehi(T + 2, 0) Else ProdTrans = ""
                
                If ProdTrans = "" And CantTanc > 0 Then
                    If CantTanc - CantTrans >= LitrosOrd Then
                    VarKbufferTancvehi(T + 1, 0) = CantTrans + LitrosOrd
                    CantTransfin = CantTrans + LitrosOrd
                        If Guardar = True Then
                        SqlUpdateTanques = " UPDATE       SATURA_VEHICULO SET SAVDEP" & Ntanque & " = " & CantTransfin & ", SAVPROD" & Ntanque & " ='" & ProdOrd & "'" & _
                                            "WHERE        (SAVFECHA = CONVERT(DATETIME,'" & FechaOrden & "',102)) AND (SAVUT = '" & Vehiculo & "') AND (SAVNVI = " & Nviaje & ")"
                        If Not DA.ExecuteQuery(SqlUpdateTanques) Then
                        End If
                        End If
                    Cabe = True
                    Else
                    ' Repartimos el producto por los tanques
                        If TanqueDedicado = "N" Then
                        LitrosRest = LitrosOrd - (CantTanc - CantTrans)
                        NtanqueRep = Ntanque
                        Do While NtanqueRep < 7 And Cabe = False
                        Cabe = SatRepartetanques(FechaOrden, Vehiculo, Nviaje, ProdOrd, VarKbufferTancvehi, T, NtanqueRep, LitrosRest, Guardar)
                        NtanqueRep = NtanqueRep + 1
                        Loop
                        If Cabe = True Then VarKbufferTancvehi(T + 1, 0) = (CantTanc - CantTrans)
                        End If
                     End If
                Else
                    If ProdTrans = ProdOrd And CantTanc - CantTrans > 0 Then
                        If CantTanc - CantTrans > LitrosOrd Then
                        VarKbufferTancvehi(T + 1, 0) = CantTrans + LitrosOrd
                        CantTransfin = CantTrans + LitrosOrd
                        If Guardar = True Then
                        SqlUpdateTanques = " UPDATE       SATURA_VEHICULO SET SAVDEP" & Ntanque & " = " & CantTransfin & ", SAVPROD" & Ntanque & " ='" & ProdOrd & "'" & _
                                            "WHERE        (SAVFECHA = CONVERT(DATETIME,'" & FechaOrden & "',102)) AND (SAVUT = '" & Vehiculo & "') AND (SAVNVI = " & Nviaje & ")"
                        If Not DA.ExecuteQuery(SqlUpdateTanques) Then
                        End If
                        End If
                        Cabe = True
                        Else
                        ' Repartimos el producto por los tanques
                        If TanqueDedicado = "N" Then
                        LitrosRest = LitrosOrd - (CantTanc - CantTrans)
                        NtanqueRep = Ntanque
                        Do While NtanqueRep < 7 And Cabe = False
                        Cabe = SatRepartetanques(FechaOrden, Vehiculo, Nviaje, ProdOrd, VarKbufferTancvehi, T, NtanqueRep, LitrosRest, Guardar)
                        NtanqueRep = NtanqueRep + 1
                        Loop
                        If Cabe = True Then VarKbufferTancvehi(T + 1, 0) = (CantTanc - CantTrans)
                        End If
                        End If
                    End If
                End If
                
                T = T + 3
                Ntanque = Ntanque + 1
                Loop
            If Cabe = False Then CabeProd = False
            P = P + 1
            Loop
            
            If CabeProd = True Then SatCHKTancaje = True Else SatCHKTancaje = False
            
    
        
'            If Cantdisp >= CantOrden Then SatCHKTancaje = True Else SatCHKTancaje = False
    
    
        End If

    End If




End Function

Public Function SatRepartetanques(ByVal FechaOrden As String, ByVal Vehiculo As String, ByVal Nviaje As Integer, ByVal ProdOrd As String, ByRef VarKbufferTancvehi As Variant, ByVal T As Integer, ByRef Ntanque As Integer, ByRef LitrosOrd As Long, ByVal Guardar As Boolean) As Boolean
Dim DA As CDataAccess
Set DA = DataConnect.DA
Dim Cabe As Boolean
Dim ProdTrans As String
Dim CantTanc As Long

Dim CantTrans As Long
Dim CantTransfin  As Long
Dim SqlUpdateTanques As String


If Not IsNull(VarKbufferTancvehi(T, 0)) Then CantTanc = VarKbufferTancvehi(T, 0) Else CantTanc = 0
If Not IsNull(VarKbufferTancvehi(T + 1, 0)) Then CantTrans = VarKbufferTancvehi(T + 1, 0) Else CantTrans = 0
If Not IsNull(VarKbufferTancvehi(T + 2, 0)) Then ProdTrans = VarKbufferTancvehi(T + 2, 0) Else ProdTrans = ""


        If ProdTrans = "" And CantTanc > 0 Then
            If CantTanc - CantTrans >= LitrosOrd Then
            VarKbufferTancvehi(T + 1, 0) = CantTrans + LitrosOrd
            CantTransfin = CantTrans + LitrosOrd
                If Guardar = True Then
                SqlUpdateTanques = " UPDATE       SATURA_VEHICULO SET SAVDEP" & Ntanque & " = " & CantTransfin & ", SAVPROD" & Ntanque & " ='" & ProdOrd & "'" & _
                                    "WHERE        (SAVFECHA = CONVERT(DATETIME,'" & FechaOrden & "',102)) AND (SAVUT = '" & Vehiculo & "') AND (SAVNVI = " & Nviaje & ")"
                If Not DA.ExecuteQuery(SqlUpdateTanques) Then
                End If
                End If
            LitrosOrd = 0
            Cabe = True
            Else
            CantTransfin = CantTanc
            VarKbufferTancvehi(T + 1, 0) = CantTanc
            LitrosOrd = LitrosOrd - CantTransfin
                If Guardar = True Then
                SqlUpdateTanques = " UPDATE       SATURA_VEHICULO SET SAVDEP" & Ntanque & " = " & CantTransfin & ", SAVPROD" & Ntanque & " ='" & ProdOrd & "'" & _
                                    "WHERE        (SAVFECHA = CONVERT(DATETIME,'" & FechaOrden & "',102)) AND (SAVUT = '" & Vehiculo & "') AND (SAVNVI = " & Nviaje & ")"
                If Not DA.ExecuteQuery(SqlUpdateTanques) Then
                End If
                End If
            End If
        Else
            If ProdTrans = ProdOrd And CantTanc - CantTrans > 0 Then
                If CantTanc - CantTrans > LitrosOrd Then
                VarKbufferTancvehi(T + 1, 0) = CantTrans + LitrosOrd
                CantTransfin = CantTrans + LitrosOrd
                If Guardar = True Then
                SqlUpdateTanques = " UPDATE       SATURA_VEHICULO SET SAVDEP" & Ntanque & " = " & CantTransfin & ", SAVPROD" & Ntanque & " ='" & ProdOrd & "'" & _
                                    "WHERE        (SAVFECHA = CONVERT(DATETIME,'" & FechaOrden & "',102)) AND (SAVUT = '" & Vehiculo & "') AND (SAVNVI = " & Nviaje & ")"
                If Not DA.ExecuteQuery(SqlUpdateTanques) Then
                End If
                End If
                LitrosOrd = 0
                Cabe = True
                Else
                ' Repartimos el producto por los tanques
                CantTransfin = CantTanc
                LitrosOrd = LitrosOrd - (CantTransfin - CantTrans)
                VarKbufferTancvehi(T + 1, 0) = CantTanc
                If Guardar = True Then
                SqlUpdateTanques = " UPDATE       SATURA_VEHICULO SET SAVDEP" & Ntanque & " = " & CantTransfin & ", SAVPROD" & Ntanque & " ='" & ProdOrd & "'" & _
                                    "WHERE        (SAVFECHA = CONVERT(DATETIME,'" & FechaOrden & "',102)) AND (SAVUT = '" & Vehiculo & "') AND (SAVNVI = " & Nviaje & ")"
                If Not DA.ExecuteQuery(SqlUpdateTanques) Then
                End If
                End If
                End If
            End If
        End If

If Cabe = True Then SatRepartetanques = True Else SatRepartetanques = False



End Function


Public Function SatUltimaDescVehi(ByRef Vehiculo As String, ByRef Fecha As String, ByRef Nviaje As Integer, ByRef Ndescarga As Integer)
Dim DA As CDataAccess
Set DA = DataConnect.DA
Dim varKBufferDescarga As Variant
Dim lngHowManyDescarga As Long

Dim SQLDescarga As String
SQLDescarga = "SELECT         MAX(SARNDESC) AS [Desc] FROM SATURACION GROUP BY SARFECHA, SARUT, SARNVI " & _
        "HAVING (SARUT = '" & Vehiculo & "') AND (SARFECHA = CONVERT(DATETIME, '" & Fecha & "', 102)) AND (SARNVI = " & Nviaje & ") "

If DA.QueryCommand(SQLDescarga, varKBufferDescarga, lngHowManyDescarga) Then
End If

If lngHowManyDescarga > 0 Then

Ndescarga = varKBufferDescarga(0, 0)

Else

Ndescarga = 0
End If

End Function

Public Function SatDatosViVehi(ByRef Vehiculo As String, ByRef Fecha As String, ByRef Nviaje As Integer, ByRef Ndescarga As Integer, ByRef Litros As Long, ByRef Longx As String, ByRef Laty As String, ByRef HoraLlegada As Date, ByRef TiempoTrabdia As Integer)
Dim DA As CDataAccess
Set DA = DataConnect.DA
Dim varKBufferDescarga As Variant
Dim lngHowManyDescarga As Long

Dim SQLDescarga As String
' Miramos los litros del viaje
SQLDescarga = "SELECT         SUM(SARLITROS) AS [Litros] FROM SATURACION GROUP BY SARFECHA, SARUT, SARNVI " & _
        "HAVING (SARUT = '" & Vehiculo & "') AND (SARFECHA = CONVERT(DATETIME, '" & Fecha & "', 102)) AND (SARNVI = " & Nviaje & ") "

If DA.QueryCommand(SQLDescarga, varKBufferDescarga, lngHowManyDescarga) Then
End If

If lngHowManyDescarga > 0 Then

Litros = varKBufferDescarga(0, 0)

Else

Litros = 0
End If

' Miramos la última posición donde se ha quedado el vehículo
If SatUltimaDescVehi(Vehiculo, Fecha, Nviaje, Ndescarga) Then
End If
SQLDescarga = "SELECT  SARLONGX, SARLATY, SARHLLEGADA  FROM SATURACION " & _
"WHERE        (SARUT = '" & Vehiculo & "') AND (SARFECHA = CONVERT(DATETIME, '" & Fecha & "', 102)) AND (SARNVI = " & Nviaje & ") AND (SARNDESC = " & Ndescarga & ")"

If DA.QueryCommand(SQLDescarga, varKBufferDescarga, lngHowManyDescarga) Then
End If

If lngHowManyDescarga > 0 Then

Longx = varKBufferDescarga(0, 0)
Laty = varKBufferDescarga(1, 0)
HoraLlegada = varKBufferDescarga(2, 0)
Else

Longx = ""
Laty = ""
End If

' Miramos las horas de trabajo del día del vehículo
SQLDescarga = "SELECT        SUM(SARTIEMPO) AS TIEMPO FROM  SATURACION GROUP BY SARFECHA, SARUT " & _
"HAVING        (SARUT = '" & Vehiculo & "') AND (SARFECHA = CONVERT(DATETIME, '" & Fecha & "', 102))"

If DA.QueryCommand(SQLDescarga, varKBufferDescarga, lngHowManyDescarga) Then
End If

If lngHowManyDescarga > 0 Then

TiempoTrabdia = varKBufferDescarga(0, 0)
Else

TiempoTrabdia = 0
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
    Set XA = New XArrayDB

    XA.ReDim 0, 5, 0, 2
    ProcesoEnCurso = False
'    XA.value(0, 0) = "Aggiornamento posizioni Oesia."
'    XA.value(0, 1) = "In attesa"
'    XA.value(0, 2) = "Esegui"
    
'    XA.value(0, 0) = "Aggiornamento posizioni Deusto."
'    XA.value(0, 1) = "En Espera"
'    XA.value(0, 2) = "Ejecutar"
    
    XA.value(0, 0) = "Envio Saturacion Cepsa"
    XA.value(0, 1) = "En Espera"
    XA.value(0, 2) = "Ejecutar"
    
    XA.value(1, 0) = "Forzar Rellenar Posiciones"
    XA.value(1, 1) = "En Espera"
    XA.value(1, 2) = "Ejecutar"

    XA.value(2, 0) = "No Utilizado"
    XA.value(2, 1) = "En Espera"
    XA.value(2, 2) = "Ejecutar"
    
    XA.value(3, 0) = "Recoger Fichero Click & Find / Startit"
    XA.value(3, 1) = "En Espera"
    XA.value(3, 2) = "Ejecutar"
    
    XA.value(4, 0) = "Envio Fichero Click & Find / Startit"
    XA.value(4, 1) = "En Espera"
    XA.value(4, 2) = "Ejecutar"

    XA.value(5, 0) = "Envio Fichero GIS CEPSA CLICK&FIND"
    XA.value(5, 1) = "En Espera"
    XA.value(5, 2) = "Ejecutar"
    
    gridProcedures.Array = XA
    gridProcedures.ReBind

    SetTimerInterval CLng(gmRegistry.GetSetting("Options", REGISTRY_KEY_NAME_EXECUTE_SECONDS, "15", True)), True
    StartTimer

    strLogLevels = gmRegistry.GetSetting("Options", REGISTRY_KEY_NAME_LOG_LEVELS, "IDEW", True)
    Me.mskFileNameEUT = gmRegistry.GetSetting("Options", REGISTRY_KEY_NAME_PATH_EUT, "", True)
    Me.mskFileNameTPH = gmRegistry.GetSetting("Options", REGISTRY_KEY_NAME_PATH_TPH, "", True)
    Me.txt_imppos = gmRegistry.GetSetting("Options", REGISTRY_KEY_NAME_PATH_POSOES, "", True)
    Me.txt_hispos = gmRegistry.GetSetting("Options", REGISTRY_KEY_NAME_PATH_HISPOES, "", True)
    Me.DIR_CLICKFIND = gmRegistry.GetSetting("Options", REGISTRY_KEY_NAME_PATH_CLICKFIND, "", True)
    MaskEdUnidad = gmRegistry.GetSetting("Options", REGISTRY_KEY_NAME_UNIDAD, "", True)
    MaskRepos = gmRegistry.GetSetting("Options", REGISTRY_KEY_NAME_REPOS, "", True)
    
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
    
    Set mobjOutFileCLICK = New CTextFile
    With mobjOutFileCLICK
        .Add "Matricula", 15, "1"
        .Add "Fechapos", 25, "2"
        .Add "Longitud", 20, "3"
        .Add "Latitud", 20, "4"
        .Add "Velocidad", 15, "5"
    End With
    
    Set mobjOutFileSTARTIT = New CTextFile
    With mobjOutFileSTARTIT
        .Add "Matricula", 15, "1"
        .Add "Fechapos", 25, "2"
        .Add "Longitud", 20, "3"
        .Add "Latitud", 20, "4"
        .Add "Velocidad", 15, "5"
        .Add "Tanque1", 2, "6"
        .Add "Tanque2", 2, "7"
        .Add "Tanque3", 2, "8"
        .Add "Tanque4", 2, "9"
        .Add "Tanque5", 2, "10"
        .Add "Tanque6", 2, "11"
    End With
    Set mobjOutFileSaturacionTot = New CTextFile
    With mobjOutFileSaturacionTot
        .Add "FechaHora", 12, "1"
        .Add "Fecha", 8, "2"
        .Add "Transportista", 4, "3"
        .Add "Centro", 4, "4"
        .Add "Ocupado", 6, "5"
        .Add "Tiempo_Disp", 6, "6"
        .Add "Ocupacion", 12, "7"
    End With
    Set mobjOutFileSaturacionTram = New CTextFile
    With mobjOutFileSaturacionTram
        .Add "FechaHora", 12, "1"
        .Add "Fecha", 8, "2"
        .Add "Transportista", 4, "3"
        .Add "Centro", 4, "4"
        .Add "Tramo", 2, "5"
        .Add "Vehiculo", 11, "6"
        .Add "MatTract", 18, "7"
        .Add "MatCist", 18, "8"
        .Add "Ocupado", 6, "9"
        .Add "Tiempo_Disp", 6, "10"
        .Add "Ocupacion", 12, "11"
    End With
    'mskLastEvent.Text = gmRegistry.GetSetting("Options", "LastEvent_Oesia", "0", True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    gmRegistry.SaveSetting "Options", REGISTRY_KEY_NAME_PATH_TPH, Me.mskFileNameTPH, True
    gmRegistry.SaveSetting "Options", REGISTRY_KEY_NAME_PATH_EUT, Me.mskFileNameEUT, True
    'gmRegistry.SaveSetting "Options", "LastEvent_Oesia", mskLastEvent.Text, True
    gmRegistry.SaveSetting "Options", REGISTRY_KEY_NAME_EXECUTE_SECONDS, Me.mskTimerSeconds, True
    gmRegistry.SaveSetting "Options", REGISTRY_KEY_NAME_PATH_POSOES, Me.txt_imppos, True
    gmRegistry.SaveSetting "Options", REGISTRY_KEY_NAME_PATH_HISPOES, Me.txt_hispos, True
    gmRegistry.SaveSetting "Options", REGISTRY_KEY_NAME_PATH_CLICKFIND, Me.DIR_CLICKFIND, True
    gmRegistry.SaveSetting "Options", REGISTRY_KEY_NAME_UNIDAD, MaskEdUnidad.Text, True
    gmRegistry.SaveSetting "Options", REGISTRY_KEY_NAME_REPOS, MaskRepos.Text, True
    
    Set XA = Nothing

    Set frmSatellitareProcedure = Nothing
End Sub
Public Function NombreGis(NLote As String) As String
Dim i As Integer
Dim x As Integer
Dim R As String

x = Len(NLote)
R = CStr(NLote)
For i = x To 9
    R = "0" + R
Next i

NombreGis = R

End Function
Public Function ImportaGis()


If (ProcesoEnCurso = False) Then
ProcesoEnCurso = True

'    --SM -JG Detección de ejecución de proccesos
Dim ret As Boolean
Dim xy As RESOIL2.CTextFile
Set xy = New CTextFile
ret = xy.OpenFile("R:\Resoil2\Alarmas\posvehi.txt", 1)
ret = xy.CloseFile()



    Dim SQL_INSGIS As String
    Dim DA As CDataAccess
    Set DA = DataConnect.DA
    Dim SQL_Pos_error As String
    Dim ArchOrigen As String
    Dim linea, Dirtemp As String
    Dim contenido, car As String
    Dim directorio(0 To 1) As String
    Dim linfich, cont, codcar As Integer
    Dim Cordx, Cordy, IdV, letra As String
    Dim x, j, pos As Integer
    Dim Valor(0 To 6) As String
    Dim FechaUTCVálida, FechaUTCUltima, FechaServ, FechaActSis As String
    j = 0
    Dim rowbuff As Variant
    Dim lngnocap As Long
    Dim rowbuffTH As Variant
    Dim lngnocapTH As Long
    Dim VarkbufferUltPOS As Variant
    Dim lngHowManyUltPOS As Long
    Dim SQL_Ult_Pos As String
    Dim MinsSinPOs As Long
    
    'Alarma de llegada de posiciones Tecnocom
    If DA.QueryCommand("SELECT TOP 1 FEUTCVAL FROM dbo.PosicionesOesia WHERE (FEUTCVAL > GETDATE() - 0.08) AND (IDVEHIC < '800')", rowbuff, lngnocap) Then

        If (lngnocap = 0) Then
            ret = xy.OpenFile("R:\Resoil2\Alarmas\POSICIONES.TXT", 1)
            ret = xy.CloseFile()
        End If
     End If

    'Alarma de llegada de posiciones Thaumat
    If DA.QueryCommand("SELECT TOP 1 FEUTCVAL FROM dbo.PosicionesOesia WHERE (FEUTCVAL > GETDATE() - 0.16) AND (IDVEHIC >= '800')", rowbuffTH, lngnocapTH) Then

        If (lngnocapTH = 0) Then
            ret = xy.OpenFile("R:\Resoil2\Alarmas\POSICIONESTH.TXT", 1)
            ret = xy.CloseFile()
        End If
     End If

    'Alarma de posiciones incoherentes
    SQL_Pos_error = "SELECT     TOP 100 PERCENT IDVEHIC, FEACT, CONVERT(float, REPLACE(Latitud, ',', '.'), 0) AS LAT, CONVERT(float, REPLACE(Longitud, ',', '.'), 0) AS LONG FROM dbo.PosicionesOesia " & _
                    "WHERE     (FEACT >= GETDATE() - 0.03) AND (CONVERT(float, REPLACE(Latitud, ',', '.'), 0) > 4) ORDER BY ID DESC "
    If DA.QueryCommand(SQL_Pos_error, rowbuff, lngnocap) Then
      
        If (lngnocap > 0) Then
            ret = xy.OpenFile("R:\Resoil2\Alarmas\POSERROR.TXT", 1)
            ret = xy.CloseFile()
        End If
     End If

    'Directorio Origen = (0,i) Directorio destino(1,i)
    directorio(0) = txt_imppos.Text
    directorio(1) = txt_hispos.Text
    
    
   LogMessage "Inicio poceso de Importacion de Posiciones: " & vbCrLf
      
      XA.value(1, 1) = "Ejecutando"
      gridProcedures.RefetchCol 1
      
DA.ExecuteCommand ("Delete From PosicionesOesia Where FEACT < (Getdate()-2)")
    
    Dirtemp = "temp\"
    
    ArchOrigen = Dir(directorio(0) + "FicheroGPS*.txt")
            Do While ArchOrigen <> ""
                FileCopy (directorio(0) + ArchOrigen), (directorio(0) + Dirtemp + ArchOrigen)
                Kill (directorio(0) + ArchOrigen)
                ArchOrigen = Dir(directorio(0) + "FicheroGPS*.txt")
            Loop
         
        ArchOrigen = Dir(directorio(0) + Dirtemp + "FicheroGPS*.txt")
        
        Do While ArchOrigen <> ""
        

        
        'Abrimos el fichero, comprobamos que exista el final de carro y seguimos.
        Open (directorio(0) + Dirtemp + ArchOrigen) For Input As #1
        contenido = Input(LOF(1), #1)
        If Len(contenido) > 0 Then
        car = Mid(contenido, Len(contenido) - 1, 1)
        codcar = ASC(car)
        If codcar = 13 Then
        linfich = 0
            For cont = 1 To Len(contenido)
                linea = ""
                codcar = 0
                Do While (codcar <> 10 And codcar <> 13) And cont <= Len(contenido)
                    car = Mid(contenido, cont, 1)
                    linea = linea + car
                    cont = cont + 1
                    codcar = ASC(car)
                Loop
            ' Aqui puedes coger los campos de la linea que necesites para guardarlos en un variant más manejable
            'Codigo Vehiculo;Fecha UTC Última Válida;Fecha UTC Última ;Fecha Servidor ;Fecha Actual Sistema;Latitud;Longitud

                Valor(j) = ""
                    Do While pos < Len(linea) - 1
                        car = Mid(linea, pos + 1, 1)
                        pos = pos + 1
                        letra = car
                            
                        If (letra <> ";") Then
                            If (j < 7) Then
                                 Valor(j) = Valor(j) + car
                            End If
                           
                        Else
                                j = j + 1
                            If (j < 7) Then
                                Valor(j) = ""
                            End If
                        End If
                    Loop
                pos = 0
                IdV = Valor(0)
                FechaUTCVálida = Format(Valor(1), "dd/mm/yyyy hh:mm:ss")
                FechaUTCUltima = Format(Valor(2), "dd/mm/yyyy hh:mm:ss")
                FechaServ = Format(Valor(3), "dd/mm/yyyy hh:mm:ss")
                FechaActSis = Format(Valor(4), "dd/mm/yyyy hh:mm:ss")
                Cordx = Mid(Trim(Valor(5)), 1, 10)
                Cordy = Mid(Trim(Valor(6)), 1, 10)
                j = 0
                    

                MinsSinPOs = DateDiff("n", FechaUTCVálida, Now())
                If MinsSinPOs < 2880 Then
                ' Miramos la fecha de la última posicón válida que tenemos de la unidad
'                    SQL_Ult_Pos = "SELECT FEUTCVAL FROM PosicionesOesia WHERE (IDVEHIC = ?) AND (FEUTCVAL = ?)"
'                    If Not DA.QueryCommand(SQL_Ult_Pos, VarkbufferUltPOS, lngHowManyUltPOS, IdV, FechaUTCVálida) Then
'                    End If
'
'                    If lngHowManyUltPOS = 0 Then
'                        SQL_INSGIS = "Insert Into PosicionesOesia (IDVEHIC, FEUTCVAL, FEUTCULT, FESERV, FEACT, Longitud, Latitud) Values (" + IdV + ",'" + FechaUTCVálida + "','" + FechaUTCUltima + "','" + FechaServ + "','" + FechaUTCVálida + "','" + cordx + "','" + cordy + "')"
'                        If Not DA.ExecuteQuery(SQL_INSGIS) Then
'                             LogMessage "Error en la Importación de fichero: " + ArchOrigen + " El ID: " + IdV & vbCrLf
'                        End If
'                    End If
                    SQL_INSGIS = "Insert Into PosicionesOesia (IDVEHIC, FEUTCVAL, FEUTCULT, FESERV, FEACT, Longitud, Latitud) Values (" + IdV + ",'" + FechaUTCVálida + "','" + FechaUTCUltima + "','" + FechaServ + "','" + FechaUTCVálida + "','" + Cordx + "','" + Cordy + "')"
                    If Not DA.ExecuteQuery(SQL_INSGIS) Then
                         LogMessage "Error en la Importación de fichero: " + ArchOrigen + " El ID: " + IdV & vbCrLf
                    End If
                    
                End If
            Next cont
         Close #1
         
          FileCopy (directorio(0) + Dirtemp + ArchOrigen), (directorio(1) + ArchOrigen)
          Kill (directorio(0) + Dirtemp + ArchOrigen)
        ArchOrigen = Dir(directorio(0) + Dirtemp + "FicheroGPS*.txt")
       Else
       Close #1
       Kill (directorio(0) + Dirtemp + ArchOrigen)
        ArchOrigen = Dir(directorio(0) + Dirtemp + "FicheroGPS*.txt")
       End If
       Else
        Close #1
        Kill (directorio(0) + Dirtemp + ArchOrigen)
        ArchOrigen = Dir(directorio(0) + Dirtemp + "FicheroGPS*.txt")
       End If
       Loop

ProcesoEnCurso = False
        LogMessage "Poceso de captura de posiciones Finalizado " & vbCrLf
        XA.value(1, 1) = "En Espera"
        gridProcedures.RefetchCol 1
End If
 
End Function

Public Function GeneraGis()
If (ProcesoEnCurso = False) Then
ProcesoEnCurso = True
Dim DA As CDataAccess

    LogMessage "Inicio poceso de Generación GIS Tecnocom." & vbCrLf


'    --SM -JG Detección de ejecución de proccesos
Dim ret As Boolean
Dim xy As RESOIL2.CTextFile
Set xy = New CTextFile
ret = xy.OpenFile(MaskEdUnidad.Text & ":\Resoil2\Alarmas\GIS.txt", 1)
ret = xy.CloseFile()



    XA.value(2, 1) = "En Ejecucion"
    gridProcedures.RefetchCol 1

    Dim KVE As CSatellitareEvento
    Dim i As Integer
    Dim varKBufferTPH As Variant
    Dim lngHowManyTPH As Long
    Dim varKBufferEUT As Variant
    Dim lngHowManyEUT As Long
    Dim SQL_GIS_recuperar_TPH As String
    Dim SQL_GIS_recuperar_EUT As String
    Dim SQL_UPDATE_GISENVIADOS As String
    Dim IDGIS As String
    SQL_UPDATE_GISENVIADOS = "UPDATE PosicionesOesia SET Enviado= 1 WHERE ID= ?"
    SQL_GIS_recuperar_TPH = "SELECT  * From GISENV WHERE  (CODTRANS = '624') And (Enviado = '0') "
    SQL_GIS_recuperar_EUT = "SELECT  * From GISENV WHERE  (CODTRANS = '004') And (Enviado = '0') "

'    Dim ERNstrTempFileName As String
'   ERNstrTempFileName = "GES_"
    Set DA = DataConnect.DA
    Set KVE = New CSatellitareEvento
    Dim NLoteTPH As String
    Dim NLoteEUT As String
    Dim FechaGL As String
    Dim HoraGL As String
    Dim NCampos As Integer
    Dim x As String
    
    
    
    With KVE
        Set .DS.DA = DataConnect.DA
        Set .DS.ErrH = DataConnect.DA.ErrH
    End With
    
   
    
    If Not DA.QueryCommand(SQL_GIS_recuperar_TPH, varKBufferTPH, lngHowManyTPH) Then
       'objLogger.logError "Query SQL_GES_recuperar."
       'GoTo ErrHandler
    End If
 
    If (lngHowManyTPH > 0) Then

            ' SELECIONAR EL NUMERO DE LOTE TPH
             With KVE.DS.DA
               NLoteTPH = .GetCounter("anaganu0", "anuprogr", "anuidanu", "030", 0)
               If NLoteTPH = -1 Then
                  'GoTo ErrHandler
               End If
               If NLoteTPH > 9999999999# Then
                  NLoteTPH = 1
               End If
             End With

'NLoteTPH = TPH3737
 If mobjOutFile.OpenFile(mskFileNameTPH.Text + "GIS_" + NombreGis(NLoteTPH) + ".DAT", 1) Then
 End If
            For i = 0 To UBound(varKBufferTPH, 2)
                mobjOutFile.Item("1").value = "0624"
                mobjOutFile.Item("2").value = NLoteTPH
                mobjOutFile.Item("3").value = Format$(Date, "YYYYMMDD")
                mobjOutFile.Item("4").value = Format$(Time, "hh:mm:ss")
                mobjOutFile.Item("5").value = lngHowManyTPH
                mobjOutFile.Item("6").value = varKBufferTPH(2, i)
                mobjOutFile.Item("7").value = varKBufferTPH(3, i)
                mobjOutFile.Item("8").value = varKBufferTPH(4, i)
                mobjOutFile.Item("9").value = varKBufferTPH(5, i)
                'mobjOutFile.Item("10").value = MATRICULA 3 (no existe en ningún caso)
                mobjOutFile.Item("11").value = "01"
                mobjOutFile.Item("12").value = varKBufferTPH(8, i)
                mobjOutFile.Item("13").value = varKBufferTPH(7, i)
                mobjOutFile.Item("14").value = varKBufferTPH(9, i)
                mobjOutFile.Item("15").value = varKBufferTPH(10, i)
                mobjOutFile.WriteLine
            Next i
            
    End If
    mobjOutFile.CloseFile
    
    If Not DA.QueryCommand(SQL_GIS_recuperar_EUT, varKBufferEUT, lngHowManyEUT) Then
       'objLogger.logError "Query SQL_GES_recuperar."
       'GoTo ErrHandler
    End If
    If lngHowManyEUT > 0 Then
                ' SELECIONAR EL NUMERO DE LOTE EUT
                  With KVE.DS.DA
                     NLoteEUT = .GetCounter("anaganu0", "anuprogr", "anuidanu", "031", 0)
                     If NLoteEUT = -1 Then
                         'GoTo ErrHandler
                     End If
                     If NLoteEUT > 9999999999# Then
                         NLoteEUT = 1
                     End If
                  End With
                  
     If mobjOutFile.OpenFile(mskFileNameEUT.Text + "GIS_" + NombreGis(NLoteEUT) + ".DAT", 1) Then
     End If
            For i = 0 To UBound(varKBufferEUT, 2)
                mobjOutFile.Item("1").value = "0004"
                mobjOutFile.Item("2").value = NLoteEUT
                mobjOutFile.Item("3").value = Format$(Date, "YYYYMMDD")
                mobjOutFile.Item("4").value = Format$(Time, "hh:mm:ss")
                mobjOutFile.Item("5").value = lngHowManyEUT
                mobjOutFile.Item("6").value = varKBufferEUT(2, i)
                mobjOutFile.Item("7").value = varKBufferEUT(3, i)
                mobjOutFile.Item("8").value = varKBufferEUT(4, i)
                mobjOutFile.Item("9").value = varKBufferEUT(5, i)
                'mobjOutFil.Item("10").value = MATRICULA 3 (no existe en ningún caso)
                mobjOutFile.Item("11").value = "01"
                mobjOutFile.Item("12").value = varKBufferEUT(8, i)
                mobjOutFile.Item("13").value = varKBufferEUT(7, i)
                mobjOutFile.Item("14").value = varKBufferEUT(9, i)
                mobjOutFile.Item("15").value = varKBufferEUT(10, i)
                mobjOutFile.WriteLine
    Next i
    End If
    mobjOutFile.CloseFile
    

   
   'Updateamos las posiciones marcandolas como enviadas.
    If (lngHowManyTPH > 0) Then
        For i = 0 To UBound(varKBufferTPH, 2)
            IDGIS = varKBufferTPH(0, i)
            If Not DA.ExecuteCommand(SQL_UPDATE_GISENVIADOS, IDGIS) Then
            End If
        Next i
    End If
    If lngHowManyEUT > 0 Then
        For i = 0 To UBound(varKBufferEUT, 2)
            IDGIS = varKBufferEUT(0, i)
            If Not DA.ExecuteCommand(SQL_UPDATE_GISENVIADOS, IDGIS) Then
            End If
        Next i
    End If
   
   LogMessage "Poceso de generacion de GIS Tecnocom Finalizado " & vbCrLf
    
    If MaskRepos.Text >= 0 Then
    Rellenapos ("TECNO")
    Rellenapos ("THMAT")
    End If
    
        XA.value(2, 1) = "En Espera"
        gridProcedures.RefetchCol 1
ProcesoEnCurso = False
End If

 End Function


'
Public Function GeneraGisClick()
If (ProcesoEnCurso = False) Then
ProcesoEnCurso = True
Dim DA As CDataAccess

    LogMessage "Inicio poceso de Generación GIS Click&Find." & vbCrLf

    XA.value(2, 1) = "En Ejecucion"
    gridProcedures.RefetchCol 1

    Dim KVE As CSatellitareEvento
    Dim i As Integer
    Dim varKBufferTPH As Variant
    Dim lngHowManyTPH As Long
    Dim varKBufferEUT As Variant
    Dim lngHowManyEUT As Long
    Dim SQL_GIS_recuperar_TPH As String
    Dim SQL_GIS_recuperar_EUT As String
    Dim SQL_UPDATE_GISENVIADOS As String
    Dim IDGIS As String
    SQL_UPDATE_GISENVIADOS = "UPDATE PosicionesClick SET Enviado= 1 WHERE ID= ?"
    SQL_GIS_recuperar_TPH = "SELECT  * From GISCLICKENV WHERE  (CODTRANS = '624') And (Enviado = '0') "
    SQL_GIS_recuperar_EUT = "SELECT  * From GISCLICKENV WHERE  (CODTRANS = '004') And (Enviado = '0') "

'    Dim ERNstrTempFileName As String
'   ERNstrTempFileName = "GES_"
    Set DA = DataConnect.DA
    Set KVE = New CSatellitareEvento
    Dim NLoteTPH As String
    Dim NLoteEUT As String
    Dim FechaGL As String
    Dim HoraGL As String
    Dim NCampos As Integer
    Dim x As String
    Dim rowbuff As Variant
    Dim lngnocap As Long
    Dim SQL_CLICK_OESIA As String
    Dim SQL_CLICK_MATERROR As String
    Dim ret As Boolean
    Dim xy As RESOIL2.CTextFile
    Set xy = New CTextFile
    ret = xy.OpenFile(MaskEdUnidad.Text & ":\Resoil2\Alarmas\GISCLICK.txt", 1)
    ret = xy.CloseFile()
        
    '    --SM -JG Detección de ejecución de proccesos
    
    ''Comprobamos que no existen viajes programados para una UDT sin GPS CLICK ni OESIA (quitar HAVING (ANAGAAM0.AAMIDDEU IS NULL) para solo click
    
'    SQL_CLICK_OESIA = " SELECT       VIAGVVT0.VVTIDAM1, ANAGAAM0.AAMIDDEU, VIAGVVT0.VVTIDAM2 " & _
'    " FROM         VIAGVVT0 INNER JOIN " & _
'    "                         ANAGAAM0 ON VIAGVVT0.VVTIDAM1 = ANAGAAM0.AAMIDAAM LEFT OUTER JOIN " & _
'    "                         Posiciones_CLick_Tractoras ON VIAGVVT0.VVTIDAM1 = Posiciones_CLick_Tractoras.MATRICULA COLLATE SQL_Latin1_General_CP1_CI_AS LEFT OUTER JOIN " & _
'    "                         Posiciones_Click_Cisternas ON VIAGVVT0.VVTIDAM2 = Posiciones_Click_Cisternas.MATRICULA COLLATE SQL_Latin1_General_CP1_CI_AS " & _
'    " WHERE        (Posiciones_Click_Cisternas.MATRICULA IS NULL) AND (VIAGVVT0.VVTIDASC = '038') AND (VIAGVVT0.VVTDTINV > GETDATE() - 1) AND " & _
'    "                         (Posiciones_CLick_Tractoras.MATRICULA IS NULL) " & _
'    "GROUP BY VIAGVVT0.VVTIDAM1, VIAGVVT0.VVTIDAM2, ANAGAAM0.AAMIDDEU " & _
'    "HAVING        (ANAGAAM0.AAMIDDEU IS NULL) "
    
    SQL_CLICK_OESIA = " SELECT       VIAGVVT0.VVTIDAM1, ANAGAAM0.AAMIDDEU, VIAGVVT0.VVTIDAM2 " & _
    " FROM         VIAGVVT0 INNER JOIN " & _
    "                         ANAGAAM0 ON VIAGVVT0.VVTIDAM1 = ANAGAAM0.AAMIDAAM LEFT OUTER JOIN " & _
    "                         Posiciones_CLick_Tractoras ON VIAGVVT0.VVTIDAM1 = Posiciones_CLick_Tractoras.MATRICULA COLLATE SQL_Latin1_General_CP1_CI_AS LEFT OUTER JOIN " & _
    "                         Posiciones_Click_Cisternas ON VIAGVVT0.VVTIDAM2 = Posiciones_Click_Cisternas.MATRICULA COLLATE SQL_Latin1_General_CP1_CI_AS " & _
    " WHERE        (Posiciones_Click_Cisternas.MATRICULA IS NULL) AND (VIAGVVT0.VVTIDASC = '038') AND (VIAGVVT0.VVTDTINV > GETDATE() - 1) AND " & _
    "                         (Posiciones_CLick_Tractoras.MATRICULA IS NULL) " & _
    "GROUP BY VIAGVVT0.VVTIDAM1, VIAGVVT0.VVTIDAM2, ANAGAAM0.AAMIDDEU "

    
    If DA.QueryCommand(SQL_CLICK_OESIA, rowbuff, lngnocap) Then
        If (lngnocap > 0) Then
            ret = xy.OpenFile(MaskEdUnidad.Text & ":\Resoil2\Alarmas\SINGPSCLICK.TXT", 1)
            ret = xy.CloseFile()
        End If
    End If
    
    '' Comprobamos errores en matriculas CLick
    
    SQL_CLICK_MATERROR = " SELECT        PosicionesClick.MATRICULA " & _
                        " FROM            PosicionesClick LEFT OUTER JOIN " & _
                        "                         ANAGAAM0 ON PosicionesClick.MATRICULA COLLATE " & _
                        " SQL_Latin1_General_CP1_CI_AS = ANAGAAM0.AAMIDAAM " & _
                        " GROUP BY PosicionesClick.MATRICULA, ANAGAAM0.AAMIDAAM " & _
                        " HAVING        (ANAGAAM0.AAMIDAAM IS NULL) "
                        
    If DA.QueryCommand(SQL_CLICK_MATERROR, rowbuff, lngnocap) Then
        If (lngnocap > 0) Then
            ret = xy.OpenFile(MaskEdUnidad.Text & ":\Resoil2\Alarmas\ERRORGPSCLICK.TXT", 1)
            ret = xy.CloseFile()
        End If
    End If
    
                            
    With KVE
        Set .DS.DA = DataConnect.DA
        Set .DS.ErrH = DataConnect.DA.ErrH
    End With
    
   
    
    If Not DA.QueryCommand(SQL_GIS_recuperar_TPH, varKBufferTPH, lngHowManyTPH) Then
       'objLogger.logError "Query SQL_GES_recuperar."
       'GoTo ErrHandler
    End If
 
    If (lngHowManyTPH > 0) Then

            ' SELECIONAR EL NUMERO DE LOTE TPH
             With KVE.DS.DA
               NLoteTPH = .GetCounter("anaganu0", "anuprogr", "anuidanu", "030", 0)
               If NLoteTPH = -1 Then
                  GoTo ErrHandler
               End If
               If NLoteTPH > 9999999999# Then
                  NLoteTPH = 1
               End If
             End With

'NLoteTPH = TPH3737
 If mobjOutFile.OpenFile(mskFileNameTPH.Text + "GIS_" + NombreGis(NLoteTPH) + ".DAT", 1) Then
 End If
            For i = 0 To UBound(varKBufferTPH, 2)
                mobjOutFile.Item("1").value = "0624"
                mobjOutFile.Item("2").value = NLoteTPH
                mobjOutFile.Item("3").value = Format$(Date, "YYYYMMDD")
                mobjOutFile.Item("4").value = Format$(Time, "hh:mm:ss")
                mobjOutFile.Item("5").value = lngHowManyTPH
                mobjOutFile.Item("6").value = varKBufferTPH(2, i)
                mobjOutFile.Item("7").value = varKBufferTPH(3, i)
                mobjOutFile.Item("8").value = varKBufferTPH(4, i)
                mobjOutFile.Item("9").value = varKBufferTPH(5, i)
                'mobjOutFile.Item("10").value = MATRICULA 3 (no existe en ningún caso)
                mobjOutFile.Item("11").value = "01"
                mobjOutFile.Item("12").value = varKBufferTPH(8, i)
                mobjOutFile.Item("13").value = varKBufferTPH(7, i)
                mobjOutFile.Item("14").value = varKBufferTPH(9, i)
                mobjOutFile.Item("15").value = varKBufferTPH(10, i)
                mobjOutFile.WriteLine
            Next i
            
    End If
    mobjOutFile.CloseFile
    
    If Not DA.QueryCommand(SQL_GIS_recuperar_EUT, varKBufferEUT, lngHowManyEUT) Then
       'objLogger.logError "Query SQL_GES_recuperar."
       'GoTo ErrHandler
    End If
    If lngHowManyEUT > 0 Then
                ' SELECIONAR EL NUMERO DE LOTE EUT
                  With KVE.DS.DA
                     NLoteEUT = .GetCounter("anaganu0", "anuprogr", "anuidanu", "031", 0)
                     If NLoteEUT = -1 Then
                         GoTo ErrHandler
                     End If
                     If NLoteEUT > 9999999999# Then
                         NLoteEUT = 1
                     End If
                  End With
                  
     If mobjOutFile.OpenFile(mskFileNameEUT.Text + "GIS_" + NombreGis(NLoteEUT) + ".DAT", 1) Then
     End If
            For i = 0 To UBound(varKBufferEUT, 2)
                mobjOutFile.Item("1").value = "3369" 'Oilogst 0004 Eutrasur
                mobjOutFile.Item("2").value = NLoteEUT
                mobjOutFile.Item("3").value = Format$(Date, "YYYYMMDD")
                mobjOutFile.Item("4").value = Format$(Time, "hh:mm:ss")
                mobjOutFile.Item("5").value = lngHowManyEUT
                mobjOutFile.Item("6").value = varKBufferEUT(2, i)
                mobjOutFile.Item("7").value = varKBufferEUT(3, i)
                mobjOutFile.Item("8").value = varKBufferEUT(4, i)
                mobjOutFile.Item("9").value = varKBufferEUT(5, i)
                'mobjOutFil.Item("10").value = MATRICULA 3 (no existe en ningún caso)
                mobjOutFile.Item("11").value = "01"
                mobjOutFile.Item("12").value = varKBufferEUT(8, i)
                mobjOutFile.Item("13").value = varKBufferEUT(7, i)
                mobjOutFile.Item("14").value = varKBufferEUT(9, i)
                mobjOutFile.Item("15").value = varKBufferEUT(10, i)
                mobjOutFile.WriteLine
    Next i
    End If
    mobjOutFile.CloseFile
    

   
   'Updateamos las posiciones marcandolas como enviadas.
    If (lngHowManyTPH > 0) Then
        For i = 0 To UBound(varKBufferTPH, 2)
            IDGIS = varKBufferTPH(0, i)
            If Not DA.ExecuteCommand(SQL_UPDATE_GISENVIADOS, IDGIS) Then
            End If
        Next i
    End If
    If lngHowManyEUT > 0 Then
        For i = 0 To UBound(varKBufferEUT, 2)
            IDGIS = varKBufferEUT(0, i)
            If Not DA.ExecuteCommand(SQL_UPDATE_GISENVIADOS, IDGIS) Then
            End If
        Next i
    End If

   LogMessage "Poceso de generacion de GIS Click&Find Finalizado " & vbCrLf
    
    If MaskRepos.Text >= 0 Then
    Rellenapos ("CLICK")
    RellenaposVac ("CLICK")
    End If
ErrHandler:
        XA.value(2, 1) = "En Espera"
        gridProcedures.RefetchCol 1
ProcesoEnCurso = False
End If

 End Function


Private Sub gridProcedures_ButtonClick(ByVal ColIndex As Integer)
    If ColIndex = 2 Then
        Select Case gridProcedures.bookmark
            Case 0
                   GeneraSaturacion
            Case 1
                    'ImportaGis
                    RellenaposVac ("CLICK")
            Case 2
                    GeneraGis
            Case 3
                    ImportaClickFind
                    ImportaStartit
            Case 4
                    ExportaClickFind
            Case 5
                    GeneraGisClick
                
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
    'Execute_ScanEvents_Oesia
    GeneraSaturacion
    'GeneraGis
    ImportaStartit
    ImportaClickFind
    ExportaClickFind
    GeneraGisClick
    
    ' Esegue la procedura di Scan Events Deusto
    'Execute_ScanEvents_Deusto
    
    
    ' Trasmette le posizioni a Cepsa
    'Execute_Trasm_Posizioni_CEPSA
   
    
'    ' Esegue la procedura di Scan Messaggi Trasmessi
'    Execute_ScanMessageSent_KNOSOS
    
'    ' Esegue la procedura di Scan Messaggi Ricevuti
'    Execute_ScanMessageReceived_KNOSOS

End Function

Private Function Execute_ScanEvents_Oesia() As Boolean
    Dim varLastID As Variant
    Dim KVE As CSatellitareEvento

    On Error GoTo ErrHandler
    
    'varLastID = mskLastEvent.Text
    
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
    
EndProc:
    'mskLastEvent.Text = varLastID
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

        XA.value(0, 1) = "In esecuzione."
        gridProcedures.RefetchCol 1
        
        If Not .ScanEvents_DEUSTO() Then
            GoTo ErrHandler
        End If
    
    End With

    XA.value(0, 1) = "In attesa."
    gridProcedures.RefetchCol 1

    Execute_ScanEvents_Deusto = True
    Set KVE = Nothing
    Exit Function

ErrHandler:
    Set KVE = Nothing
    
    XA.value(0, 1) = "In errore."
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
    'XA.value(2, 1) = "In errore."
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



