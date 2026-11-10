VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmOCTExportEA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "COMUNICAZIONE ORDINI EA RICEVUTI"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6840
   ScaleWidth      =   9270
   Begin VB.Frame Frame4 
      Height          =   1935
      Left            =   0
      TabIndex        =   24
      Top             =   3480
      Width           =   5295
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   5
         Left            =   240
         Picture         =   "frmOCTExportEA.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1515
         Width           =   240
      End
      Begin VB.CommandButton Command6 
         Caption         =   "F7 - Apri Ordine"
         Height          =   375
         Left            =   120
         Picture         =   "frmOCTExportEA.frx":0102
         TabIndex        =   11
         Tag             =   "EDITLINK"
         Top             =   1440
         Width           =   2175
      End
      Begin TrueDBGrid60.TDBGrid tdbLog 
         Height          =   1095
         Left            =   120
         OleObjectBlob   =   "frmOCTExportEA.frx":0204
         TabIndex        =   10
         Top             =   240
         Width           =   5055
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Parametri di selezione"
      Height          =   1455
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton cmdCalendar 
         Height          =   300
         Index           =   0
         Left            =   2640
         Picture         =   "frmOCTExportEA.frx":2C08
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   375
      End
      Begin MSMask.MaskEdBox mskFileName 
         Height          =   255
         Left            =   1320
         TabIndex        =   8
         Top             =   1080
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskField1 
         Height          =   255
         Left            =   1320
         TabIndex        =   4
         Tag             =   "44"
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         Format          =   "0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox4 
         Height          =   255
         Left            =   4080
         TabIndex        =   6
         Tag             =   "45"
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   3
         Format          =   "0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDateFrom 
         Height          =   255
         Left            =   1320
         TabIndex        =   1
         Tag             =   "3"
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Data Ric."
         Height          =   255
         Index           =   12
         Left            =   360
         TabIndex        =   0
         Top             =   360
         Width           =   855
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   12
         Left            =   3120
         Tag             =   "25"
         Top             =   360
         Width           =   240
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   2
         Left            =   2640
         Tag             =   "44"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "D&al N. Lotto"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   975
      End
      Begin VB.Image imgState 
         Height          =   240
         Index           =   1
         Left            =   5400
         Tag             =   "45"
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Al N. L&otto"
         Height          =   255
         Index           =   2
         Left            =   3120
         TabIndex        =   5
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "&Percorso"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   5400
      TabIndex        =   16
      Top             =   3480
      Width           =   3855
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   7
         Left            =   2400
         TabIndex        =   27
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         Caption         =   "Ordini esclusi:"
         Height          =   255
         Index           =   6
         Left            =   1080
         TabIndex        =   26
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         Caption         =   "Ordini trasferiti correttamente:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   22
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   21
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         Caption         =   "Errori riscontrati:"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   20
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   2400
         TabIndex        =   19
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         Caption         =   "Ordini da trasferire:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblLog 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   5
         Left            =   2400
         TabIndex        =   17
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   855
      Left            =   0
      TabIndex        =   14
      Top             =   5520
      Width           =   9255
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   6
         Left            =   4560
         Picture         =   "frmOCTExportEA.frx":2D0A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "&Trasferisci"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4440
         Picture         =   "frmOCTExportEA.frx":2E0C
         TabIndex        =   30
         Tag             =   "EXPORT"
         Top             =   360
         Width           =   2175
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   7
         Left            =   240
         Picture         =   "frmOCTExportEA.frx":2F0E
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   2400
         Picture         =   "frmOCTExportEA.frx":3010
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton cmdOpenFile 
         Caption         =   "&Seleziona File..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   2280
         Picture         =   "frmOCTExportEA.frx":3112
         TabIndex        =   32
         Tag             =   "OPEN"
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton cmdReadOrders 
         Caption         =   "&Leggi Ordini"
         Height          =   375
         Left            =   120
         Picture         =   "frmOCTExportEA.frx":3214
         TabIndex        =   33
         Tag             =   "BROWSE"
         Top             =   360
         Width           =   2175
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   7560
         Picture         =   "frmOCTExportEA.frx":3316
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   435
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   7440
         Picture         =   "frmOCTExportEA.frx":3418
         TabIndex        =   12
         Tag             =   "EXIT"
         Top             =   360
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   6465
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   1773
            MinWidth        =   1764
            Text            =   "RES/OIL.2"
            TextSave        =   "RES/OIL.2"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10901
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "12.22"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "08/05/2007"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin TrueDBGrid60.TDBGrid TDBGrid1 
      Height          =   1935
      Left            =   120
      OleObjectBlob   =   "frmOCTExportEA.frx":351A
      TabIndex        =   9
      Tag             =   "SELECT"
      Top             =   1560
      Width           =   9015
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8400
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOCTExportEA.frx":DAB0
            Key             =   "GreenCheck"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOCTExportEA.frx":DBC2
            Key             =   "RedCross"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOCTExportEA.frx":DCD4
            Key             =   "Editing"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOCTExportEA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COLS = 19
Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1

Private mobjOutFile As CTextFile
Private mobjOutFileVeicoli As CTextFile

Public GridArraySelect As ArrayGridCls
Public XA As XArrayDB
Public Loading As Boolean

Public DataClass As COrdineSelezione
Public State As FormState

Private Soc As CSocieta
Private AZO As CZona
Private AXX As CParametro
Private ADZ As CDestinazione
Private OCR As COrdineRiga
Private AUT As CUnitaDiTrasporto
Private APR As CProdotto


Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    Dim LinkForm As Form
    
    Select Case Cmd
        Case fcBrowse
            If SelectRecords Then
                cmdOpenFile.Enabled = True
                If Len(mskFileName) = 0 Then
                    cmdOpenFile.Enabled = False
                End If
                cmdExport.Enabled = True
            Else
                cmdOpenFile.Enabled = False
                cmdExport.Enabled = False
            End If
        Case fcExport
            ExportRecords2
            cmdExport.Enabled = False
        Case fcOpen
            With CommonDialog1
                .fileName = mskFileName
                .DialogTitle = "Seleziona il file per Mover... "
                .InitDir = App.Path
                .Filter = "Documenti di Testo (*.txt)|*.txt|Tutti i files (*.*)|*.*| "
                .FilterIndex = 1
                .Flags = cdlOFNExplorer Or cdlOFNHideReadOnly
                .ShowSave
                mskFileName = .fileName
            End With
        Case fcOpenLinkEdit
            Set LinkForm = DisplayForm(fcORDIOCT)
            If Not LinkForm Is Nothing Then
                With LinkForm
                    Set .ParentForm = Me
                    .GetRecord GetSelectedID(tdbLog)
                    .SetState fsEdit
                    Me.Enabled = False
                End With
            End If
        Case fcExit
            Unload Me
    End Select
End Sub

' --- Esportazione ordini in formato AGIP
Private Sub ExportRecords2()
    On Error GoTo ErrHandler
    
    Dim lngRow As Long
    Dim lngNumRows As Long
    Dim lngQta As Long
    Dim lngCountOK As Long
    Dim lngCountFail As Long
    Dim lngCountToGo As Long
    Dim lngCountExcluded As Long
    Dim lngCountLotto As Long
    Dim intRowsOrder As Integer
    Dim i As Long
    Dim j As Long
    Dim intIndexProdotto As Integer
    Dim strIDOrdine As String
    Dim strIDProdotto As String
    Dim strIDParam As String
    Dim strIDZona As String
    Dim strNota3 As String
    Dim strRowState As String
    Dim strFileName As String
    Dim strNumLotto As String
    Dim strCondTecnico As String
    Dim strErrorMessage As String
    Dim strHeader As String
    Dim intRowError As Errors
    Dim varProdotti(1 To 10) As String
    Dim blnFlagProdottiOK As Boolean
    Dim blnFound As Boolean
    Dim blnGC As Boolean
    Dim ErrCode As Errors
    Dim C As CError
    
    Me.MousePointer = vbHourglass
    Me.Enabled = False
    ErrCode = NoError
    ' --- Leggo i valori di default del sistema
    strIDParam = ID_PARAM
    ' ---
    
    
    strNumLotto = vbNullString
    lngNumRows = GridArraySelect.RowCount - 1
    lngCountToGo = lngNumRows + 1
    
    ' --- Ciclo di elaborazione degli ordini
    For lngRow = 0 To lngNumRows
        intRowError = NoError
        
        ' Se ci troviamo all'inizio di un nuovo lotto apriamo un nuovo file
        If strNumLotto <> GridArraySelect.value(lngRow, 0) Then
        
            ' Se non è il primo file chiudiamo il file precedente
            If Len(strNumLotto) > 0 Then
                mobjOutFile.CloseFile
            End If
            
            strNumLotto = GridArraySelect.value(lngRow, 0)
            lngCountLotto = 0
            For j = lngRow To lngNumRows
                If GridArraySelect.value(j, 0) <> strNumLotto Then
                    Exit For
                End If
                lngCountLotto = lngCountLotto + 1
            Next j
            
            strFileName = mskFileName & "\ECOPED_002881_" & Format(strNumLotto, "0000") & ".LIS"
            If Not mobjOutFile.OpenFile(strFileName, 1) Then
                ErrCode = FileNotFound
                GoTo ErrHandler
            End If
            strHeader = Format(lngCountLotto, "0000") & "002881" & Format(GridArraySelect.value(lngRow, 8), "yyyymmdd") & Format(GridArraySelect.value(lngRow, 9), "hhnn")
            Print #mobjOutFile.FileNum, strHeader
        End If
        
        With mobjOutFile
            .ClearFields
            If Left(GridArraySelect.value(lngRow, 1), 2) = "C0" Then
                .Item("1").value = Mid(GridArraySelect.value(lngRow, 1), 3)
            Else
                .Item("1").value = GridArraySelect.value(lngRow, 1)
            End If
            .Item("2").value = Mid(GridArraySelect.value(lngRow, 2), 3)
            .Item("3").value = GridArraySelect.value(lngRow, 3)
            .Item("6").value = Format(GridArraySelect.value(lngRow, 4), "yyyymmdd")
            .Item("7").value = GridArraySelect.value(lngRow, 5)
            .Item("8").value = GridArraySelect.value(lngRow, 6)
        End With
        
        strIDOrdine = GridArraySelect.value(lngRow, 7)
        OCR.DS.Fields(OCR.DS.KeyFields(1)).FindValue = strIDOrdine
        
        If Not OCR.GetList(True) Then
            intRowError = RowsNotFound
            GoTo LogResult
        End If
        intRowsOrder = OCR.DS.HowMany - 1
        
        For i = 0 To intRowsOrder
            strIDProdotto = OCR.DS.ListItem(1, i)
            If IsEmpty(OCR.DS.ListItem(4, i)) Then
                intRowError = MissingQuantity
                GoTo LogResult
            End If
            APR.DS.Field("Codice").FindValue = strIDProdotto
            APR.Find
            strIDProdotto = APR.DS.MultiSoc.DS.GetFirmCode(ID_SOC_CEPSA, 3)
            
            'APR.DecodeFirmCode(ID_SOC_CEPSA, strIDProdotto, True)
            lngQta = CLng(OCR.DS.ListItem(4, i))
            If lngQta = 0 Then
                intRowError = MissingQuantity
                GoTo LogResult
            End If
            intIndexProdotto = 1
            With mobjOutFile
                .Item("4").value = strIDProdotto
                .Item("5").value = Format$(lngQta, "000000")
                If Not .WriteLine(True) Then
                    intRowError = OutputError
                    GoTo LogResult
                End If
             End With
        Next i
        
        
        
LogResult:
        ' --- Scrivo il risultato dell'elaborazione dell'ordine
        If intRowError = NoError Then
            lngCountOK = lngCountOK + 1
        Else
            If intRowError = ZonaEsclusaExport Then
                lngCountExcluded = lngCountExcluded + 1
            Else
                lngCountFail = lngCountFail + 1
            End If
        End If
        lngCountToGo = lngCountToGo - 1
        lblLog(2) = lngCountOK
        lblLog(4) = lngCountFail
        lblLog(5) = lngCountToGo
        lblLog(7) = lngCountExcluded
        Select Case intRowError
            Case NoError
                strErrorMessage = "OK"
            Case OutputError
                strErrorMessage = "Errore durante la scrittura sul file"
            Case NoProductInParam
                strErrorMessage = "Prodotto non trovato tra i parametri di sistema"
            Case RowsNotFound
                strErrorMessage = "Errore durante la lettura delle righe di ordine"
            Case NoSetState
                strErrorMessage = "Impossibile cambiare lo stato dell'ordine"
            Case MissingQuantity
                strErrorMessage = "Quantità mancante in una riga di ordine"
            Case NoZona
                strErrorMessage = "Codice Zona non trovato"
            Case ZonaEsclusaExport
                strErrorMessage = "Zona esclusa dal trasferimento a Mover"
            Case BadIDOrderSoc
                strErrorMessage = "L'ordine contiene prodotti duplicati"
        End Select
        XA.value(lngRow, 1) = strErrorMessage
        With tdbLog
            .Refresh
            .bookmark = lngRow
        End With
        TDBGrid1.bookmark = lngRow
        ' ---
        
        DoEvents
        
    Next lngRow
    ' ---
    
    GoTo EndProc
ErrHandler:
    Dim strMessage As String
    Dim strInfo As String
    Dim lngNativeCode As Long
    
    Select Case ErrCode
        Case FileNotFound
            strMessage = "Impossibile aprire il file."
            strInfo = "Nome del file: " & strFileName
        Case InvalidHeader
            strMessage = "Impossibile trovare il parametro di sistema."
            strInfo = "Codice Parametro: " & strIDParam
        Case MissingDefaults
            strMessage = "Parametri di sistema mancanti."
            strInfo = "Default mancanti: Prodotti"
        Case NoError
            Select Case Err.Number
                Case Else
                    ErrCode = UnknownError
                    strMessage = "Errore Sconosciuto."
                    strInfo = Err.Description
                    lngNativeCode = Err.Number
            End Select
    End Select
    Set C = ErrH.SetError(ErrCode, "frmOCTExportMover", _
                "ExportRecords", lngNativeCode, strMessage, 10, strInfo)
    DisplayError ErrH.LastError
EndProc:
    mobjOutFile.CloseFile
    Me.MousePointer = vbDefault
    Me.Enabled = True
    Exit Sub

End Sub

Private Function GetSelectedID(DBGrid As TDBGrid) As Variant
    If IsNull(DBGrid.bookmark) Then
        GetSelectedID = vbNullString
    Else
        GetSelectedID = XA.value(DBGrid.bookmark, 0)
    End If
End Function

Private Function SelectRecords() As Boolean
    On Error GoTo ErrHandler
    
    Me.MousePointer = vbHourglass
    If DataClass.DS.Fields.Validate Then
        If DataClass.DS.ScanClientCond Then
            If Not DataClass.GetSelectionExportEA() Then
                GoTo ErrHandler
            End If
            DisplaySelection
            GoTo EndProc
        End If
    Else
        ShowValidationError DataClass.DS
        GoTo ErrHandler
    End If
    GoTo EndProc
ErrHandler:
    Me.MousePointer = vbDefault
    SelectRecords = False
    Exit Function
EndProc:
    Me.MousePointer = vbDefault
    SelectRecords = True
End Function

Private Sub DisplaySelection()
    Dim varLog As Variant
    Dim i As Long
    Dim lngRows As Long
    
    lngRows = DataClass.DS.HowMany
    RefreshUnboundGrid TDBGrid1, DataClass.DS.List, _
                       GridArraySelect, lngRows, MAX_COLS

    lblLog(2) = 0
    lblLog(4) = 0
    lblLog(5) = lngRows
    lblLog(7) = 0
    If lngRows > 0 Then
    ReDim varLog(0 To 1, 0 To lngRows - 1)
    With DataClass.DS
        For i = 0 To lngRows - 1
            varLog(0, i) = .ListItem(7, i)
        Next i
    End With
    RefreshXArray XA, tdbLog, varLog
    Else
        ClearXArray XA, tdbLog
    End If
End Sub





Private Sub cmdCalendar_Click(Index As Integer)
    Dim strTitle As String
    Dim M As MaskEdBox
    
    Select Case Index
        Case 0
            Set M = mskDateFrom
            strTitle = "Data ricezione"
    End Select
    GetCalendarDate M, strTitle
End Sub

Private Sub ErrH_OnError()
    DisplayError ErrH.LastError
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    ClientKeyUp KeyCode, Me
End Sub

Private Sub Form_Load()
    Dim D As Variant
    
    Loading = True
    Set GridArraySelect = New ArrayGridCls
    Set XA = New XArrayDB
    
    mskFileName = gmRegistry.GetSetting("Options", "OrderExportPathEA", , True)
    
    Set ErrH = New CErrorHandler
    
    Set DataClass = New COrdineSelezione
    With DataClass.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        Set .FormHandle = Me
        .FormType = ftEdtRcd1
        .Fields("3").Operator = " = "
        .Fields("44").MaxLength = 4
        .Fields("45").MaxLength = 4
        .ConfigureInterface
        ConfigureObjects Me
    End With
    
    Set Soc = New CSocieta
    Set Soc.DS.DA = DataConnect.DA
    SetValueItems TDBGrid1.Columns(1), Soc
    
    Set AXX = New CParametro
    With AXX.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set OCR = New COrdineRiga
    With OCR.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set ADZ = New CDestinazione
    With ADZ.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        'Set .MultiSoc.DS.DA = DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set APR = New CProdotto
    With APR.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set mobjOutFile = New CTextFile
    With mobjOutFile
        .Add "Numero Ordine", 11, "1"
        .Add "Codice Destinazione", 8, "2"
        .Add "Ragione Sociale", 20, "3"
        .Add "Prodotto", 7, "4", , fftNumerico
        .Add "Quantità", 9, "5", , fftNumerico
        .Add "DataConsegna", 8, "6"
        .Add "Note 1", 38, "7"
        .Add "Filler 1", 40, "8"
    End With
    Loading = False
End Sub


Private Sub Form_Unload(Cancel As Integer)
    gmRegistry.SaveSetting "Options", "OrderExportPathEA", mskFileName, True
    Set XA = Nothing
    Set GridArraySelect = Nothing
    Set AXX = Nothing
    Set ADZ = Nothing
    Set OCR = Nothing
    Set AZO = Nothing
    Set AUT = Nothing
    Set APR = Nothing
    Set DataClass = Nothing
    Set Soc = Nothing
    Set mobjOutFile = Nothing
    Set mobjOutFileVeicoli = Nothing
    Set frmOCTExportEA = Nothing
End Sub


Private Sub mskFileName_Change()
    cmdExport.Enabled = (Len(mskFileName) > 0)
End Sub


