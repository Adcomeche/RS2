VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmOCTImport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RICEZIONE ORDINI"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10350
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6915
   ScaleWidth      =   10350
   Begin VB.Frame Frame2 
      Caption         =   "Comandi Disponibili"
      Height          =   735
      Left            =   0
      TabIndex        =   24
      Top             =   5760
      Width           =   10335
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   8640
         Picture         =   "frmOCTImport.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   315
         Width           =   240
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F3 - Esci"
         Height          =   375
         Left            =   8520
         Picture         =   "frmOCTImport.frx":0102
         TabIndex        =   23
         Tag             =   "EXIT"
         Top             =   240
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   6540
      Width           =   10350
      _ExtentX        =   18256
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
            Object.Width           =   12806
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "14.54"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "24/07/2006"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   5655
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   10335
      Begin VB.Frame Frame3 
         Height          =   3255
         Left            =   6360
         TabIndex        =   13
         Top             =   120
         Width           =   3855
         Begin VB.Label lblLog 
            Caption         =   "Ordini esclusi:"
            Height          =   255
            Index           =   7
            Left            =   1560
            TabIndex        =   34
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   6
            Left            =   1560
            TabIndex        =   33
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   5
            Left            =   1560
            TabIndex        =   32
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Caption         =   "Ordini da importare:"
            Height          =   255
            Index           =   3
            Left            =   1560
            TabIndex        =   31
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   4
            Left            =   1560
            TabIndex        =   30
            Top             =   2880
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Caption         =   "Errori riscontrati:"
            Height          =   255
            Index           =   1
            Left            =   1560
            TabIndex        =   29
            Top             =   2640
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   2
            Left            =   1560
            TabIndex        =   28
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label lblLog 
            Caption         =   "Ordini importati correttamente:"
            Height          =   255
            Index           =   0
            Left            =   1560
            TabIndex        =   27
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label lblField 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   21
            Top             =   2280
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Ora Invio"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   20
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label lblField 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   19
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Data Invio"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   18
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label lblField 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   17
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Progressivo Invio"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   16
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblField 
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   15
            Top             =   480
            Width           =   3615
         End
         Begin VB.Label Label1 
            Caption         =   "Ente Emittente"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1935
         Left            =   120
         TabIndex        =   26
         Top             =   1440
         Width           =   6135
         Begin TrueDBGrid60.TDBGrid tdbLog 
            Height          =   1575
            Left            =   120
            OleObjectBlob   =   "frmOCTImport.frx":0204
            TabIndex        =   37
            Top             =   240
            Width           =   5895
         End
      End
      Begin VB.Frame Frame4 
         Height          =   2175
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   3360
         Width           =   10095
         Begin TrueDBGrid60.TDBGrid TDBGrid1 
            Height          =   1815
            Left            =   120
            OleObjectBlob   =   "frmOCTImport.frx":6160
            TabIndex        =   7
            Top             =   240
            Width           =   9855
         End
      End
      Begin VB.Frame Frame4 
         Height          =   2175
         Index           =   1
         Left            =   120
         TabIndex        =   35
         Top             =   3360
         Visible         =   0   'False
         Width           =   10095
         Begin TrueDBGrid60.TDBGrid TDBGrid2 
            Height          =   1815
            Left            =   120
            OleObjectBlob   =   "frmOCTImport.frx":131EE
            TabIndex        =   36
            Top             =   240
            Width           =   9855
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   6
         Left            =   3120
         Picture         =   "frmOCTImport.frx":1AF94
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1110
         Width           =   240
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "&Importa"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3000
         Picture         =   "frmOCTImport.frx":1B4C6
         TabIndex        =   2
         Tag             =   "IMPORT"
         Top             =   1035
         Width           =   1455
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   7
         Left            =   1680
         Picture         =   "frmOCTImport.frx":1B5C8
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1110
         Width           =   240
      End
      Begin VB.CommandButton cmdReadFile 
         Caption         =   "&Leggi"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         Picture         =   "frmOCTImport.frx":1BAFA
         TabIndex        =   1
         Tag             =   "OPEN"
         Top             =   1035
         Width           =   1455
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   5
         Left            =   240
         Picture         =   "frmOCTImport.frx":1BBFC
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1110
         Width           =   240
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Sfoglia..."
         Height          =   375
         Left            =   120
         Picture         =   "frmOCTImport.frx":1BCFE
         TabIndex        =   0
         Tag             =   "BROWSE"
         Top             =   1035
         Width           =   1455
      End
      Begin MSMask.MaskEdBox mskFileName 
         Height          =   255
         Left            =   1200
         TabIndex        =   4
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   5760
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSMask.MaskEdBox mskFileNameDettaglio 
         Height          =   255
         Left            =   1200
         TabIndex        =   6
         Top             =   600
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   25
         PromptChar      =   "_"
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   8400
         Top             =   1680
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
               Picture         =   "frmOCTImport.frx":1BE00
               Key             =   "GreenCheck"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOCTImport.frx":1BF12
               Key             =   "RedCross"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOCTImport.frx":1C024
               Key             =   "Editing"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "File &Dettaglio"
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "File &Ordini"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmOCTImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_COLS = 3
Private Const COL_COD_ORD_LOG = 1
Private Const COL_COD_PROD_LOG = 4
Private Const COL_RESULT_LOG = 10

Private XA As XArrayDB
Private XALog As XArrayDB
Private AXX As CParametro
Private ADZ As CDestinazione
Private OCT As COrdine
Private OCR As COrdineRiga
Private APR As CProdotto
Private APD As CPosDoganale
Private ADE As CDeposito
Private AEE As CEnteEmittente
Private ATV As CTipoVendita
Private AZO As CZona

Private mobjInFile As CTextFile
Private mobjInFileDettaglio As CTextFile
Private mstrIDEnte As String
Private mintFileInput As Integer
Private mstrUltimoProgressivo As String

Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean

Private WithEvents ErrH As CErrorHandler
Attribute ErrH.VB_VarHelpID = -1


Private Function ConvertDate001(strDate As String) As Date

    If strDate = vbNullString Then
        ConvertDate001 = TimeSerial(0, 0, 0)
    Else
        ConvertDate001 = TimeSerial(CLng(Left(strDate, 2)), CLng(Right(strDate, 2)), 0)
    End If

End Function

Private Function CheckDestinazione001(ByRef strIDDestinazione As String, _
                                      ByVal strCodDest As String, _
                                      ByVal strIDClienteDefault As String, _
                                      ByVal strIDDeposito As String, _
                                      ByRef strNewLocalita As String) As Long
    Dim intRowError As Errors
    Dim strLocalita As String
    Dim strNewZona As String
    Dim strNewIndirizzo As String
    Dim blnConcurErr As Boolean
    Dim lngKmRaggio As Long
    Dim dateOrarioApertura1 As Date
    Dim dateOrarioChiusura1 As Date
    Dim dateOrarioApertura2 As Date
    Dim dateOrarioChiusura2 As Date
    Dim dateOrarioApertura3 As Date
    Dim dateOrarioChiusura3 As Date
   
    intRowError = NoError
    ' --- Cerco la destinazione
    strIDDestinazione = ADZ.DecodeFirmCode(ID_SOC, strCodDest, True)
    
    ' --- La destinazione è già stata inserita ?
    If strIDDestinazione = vbNullString Then
        ' --- No : devo inserirla ora
        With ADZ.DS.Fields
            RecordToADZ ADZ, mobjInFile, strIDClienteDefault, strIDDeposito
            If Not ADZ.Insert Then
                intRowError = DestInsertError
                GoTo ExitFunction
            End If
            If Not ADZ.DS.MultiSoc.Insert Then
                intRowError = DestSocInsertError
                GoTo ExitFunction
            End If
        End With
        strIDDestinazione = strCodDest
    Else
        ' --- Controllo che la località non sia cambiata
        ADZ.DS.Fields("0").FindValue = strIDDestinazione
        If Not ADZ.Find Then
            intRowError = NoDestinazione
            GoTo ExitFunction
        End If
        strLocalita = UCase(Trim$(ADZ.DS.Fields("5").value))
        strNewLocalita = ADZ.DS.DA.MakeSQL(mobjInFile("7").value, 0) 'dtAlfa)
        strNewLocalita = Mid$(strNewLocalita, 2, Len(strNewLocalita) - 2)
        
        dateOrarioApertura1 = ConvertDate001(mobjInFile("26").value)
        dateOrarioChiusura1 = ConvertDate001(mobjInFile("27").value)
        dateOrarioApertura2 = ConvertDate001(mobjInFile("28").value)
        dateOrarioChiusura2 = ConvertDate001(mobjInFile("29").value)
        dateOrarioApertura3 = ConvertDate001(mobjInFile("30").value)
        dateOrarioChiusura3 = ConvertDate001(mobjInFile("31").value)
        
        If strLocalita <> strNewLocalita Or _
            dateOrarioApertura1 <> TimeValue(ADZ.DS.Fields("33").value) Or _
            dateOrarioChiusura1 <> TimeValue(ADZ.DS.Fields("34").value) Or _
            dateOrarioApertura2 <> TimeValue(ADZ.DS.Fields("35").value) Or _
            dateOrarioChiusura2 <> TimeValue(ADZ.DS.Fields("36").value) Or _
            dateOrarioApertura3 <> TimeValue(ADZ.DS.Fields("37").value) Or _
            dateOrarioChiusura3 <> TimeValue(ADZ.DS.Fields("38").value) Then
            
            ' --- Se la località è cambiata aggiorno la destinazione
            strNewZona = mobjInFile("2").value
            strNewIndirizzo = mobjInFile("6").value
            If Len(mobjInFile("15").value) > 0 Then
                lngKmRaggio = mobjInFile("15").value
            End If
            With ADZ.DS
                .Fields("4").NewValue = strNewIndirizzo
                .Fields("4").Updated = True
                .Fields("5").NewValue = strNewLocalita
                .Fields("5").Updated = True
                .Fields("15").NewValue = strIDDeposito
                .Fields("15").Updated = True
                .Fields("16").NewValue = strNewZona
                .Fields("16").Updated = True
                .Fields("29").NewValue = lngKmRaggio
                .Fields("29").Updated = True
                .Fields("33").NewValue = dateOrarioApertura1
                .Fields("33").Updated = True
                .Fields("34").NewValue = dateOrarioChiusura1
                .Fields("34").Updated = True
                .Fields("35").NewValue = dateOrarioApertura2
                .Fields("35").Updated = True
                .Fields("36").NewValue = dateOrarioChiusura2
                .Fields("36").Updated = True
                .Fields("37").NewValue = dateOrarioApertura3
                .Fields("37").Updated = True
                .Fields("38").NewValue = dateOrarioChiusura3
                .Fields("38").Updated = True
                ADZ.BeginSave blnConcurErr
                .DA.EndTrans ETCOMMIT
                .EndSave
            End With
            ' ---
            intRowError = DestHasChanged
            GoTo ExitFunction
        Else
            If Len(ADZ.DS.Fields("29").value) > 0 Then
                lngKmRaggio = ADZ.DS.Fields("29").value
            Else
                lngKmRaggio = 0
            End If
            If lngKmRaggio <> mobjInFile("15").value Then
                With ADZ.DS
                    ' Km Raggio
                    .Fields("29").NewValue = mobjInFile("15").value
                    .Fields("29").Updated = True
                    ' Deposito
                    .Fields("15").NewValue = strIDDeposito
                    .Fields("15").Updated = True
                    ADZ.BeginSave blnConcurErr
                    .DA.EndTrans ETCOMMIT
                    .EndSave
                End With
            End If
        End If
        ' ---
    End If
    ' ---
ExitFunction:
    CheckDestinazione001 = intRowError
End Function


Private Function CheckOrderRow001(ByRef strIDDeposito As String, _
                                  ByRef strIDEnteEmittente As String, _
                                  ByRef strIDProdotto As String, _
                                  ByRef strIDTipoVendita As String, _
                                  ByRef strIDZona As String, _
                                  ByRef strCodDest As String) As Long
    Dim intRowError As Errors
    
    intRowError = NoError
    
    ' --- Controllo che esistano Deposito, Zona, Tipo Vendita, Ente Emittente, Prodotto
    strIDDeposito = ADE.DecodeFirmCode2(ID_SOC, mobjInFile("13").value, True)
    strIDEnteEmittente = AEE.DecodeFirmCode(ID_SOC, mstrIDEnte, True)
    strIDProdotto = APR.DecodeFirmCode(ID_SOC, mobjInFile("17").value, True)

    If mobjInFile("19").value = vbNullString Then
        intRowError = NoTipoVendita
        GoTo ExitFunction
    End If
    strIDTipoVendita = ATV.DecodeFirmCode(ID_SOC, mobjInFile("19").value, True)
    If strIDTipoVendita = vbNullString Then
        intRowError = NoTipoVendita
        GoTo ExitFunction
    End If
    strIDZona = mobjInFile("2").value
    AZO.DS.Fields("Codice").FindValue = strIDZona
    If Not AZO.Find Then
        intRowError = NoZona
        GoTo ExitFunction
    End If
    If AZO.DS.HowMany = 0 Then
        intRowError = NoZona
        GoTo ExitFunction
    End If
    ' --- Controllo che la zona non sia da escludere
    ' --- dall'importazione
    If AZO.DS.Fields("2").value <> 0 Then
        intRowError = ZonaEsclusaImport
        GoTo ExitFunction
    End If
    ' ---
    
    If strIDDeposito = vbNullString Then
        intRowError = NoDeposito
        GoTo ExitFunction
    End If
    If strIDEnteEmittente = vbNullString Then
        intRowError = NoEnteEmittente
        GoTo ExitFunction
    End If
    If strIDProdotto = vbNullString Then
        intRowError = NoProdotto
        GoTo ExitFunction
    End If
    ' ---
    
    strCodDest = Format$(mobjInFile("4").value, "00000000")
    If strCodDest = vbNullString Then
        intRowError = NoDestinazione
        GoTo ExitFunction
    End If
    
    
ExitFunction:
    CheckOrderRow001 = intRowError
End Function


Public Sub DoCommand(Cmd As FormCommand)
    Dim blnResult As Boolean
    
    Select Case Cmd
        Case fcBrowse
            With CommonDialog1
                .fileName = mskFileName
'                .InitDir = App.Path
                .FilterIndex = 1
                .Flags = cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNNoChangeDir
                .Filter = "Documenti di Testo (*.txt)|*.txt|Tutti i files (*.*)|*.*| "
'                Select Case gmstrCodInstallazione
'                    Case ciInst1
                #If VersionID = 1 Then
                        .DialogTitle = "Seleziona il file degli ordini... "
'                    Case ciInst2
                #ElseIf (VersionID = 2) Then 'Or (VersionID = 5) Then
                        Select Case mintFileInput
                            Case 1
                                .DialogTitle = "Seleziona il file degli ordini... "
                            Case 2
                                .DialogTitle = "Seleziona il file di dettaglio... "
                        End Select
'                End Select
                #End If
                
                .ShowOpen
'                Select Case gmstrCodInstallazione
'                    Case ciInst1
                #If VersionID = 1 Then
                        mskFileName = .fileName
'                    Case ciInst2
                #ElseIf (VersionID = 2) Then 'Or (VersionID = 5) Then
                        Select Case mintFileInput
                            Case 1
                                mskFileName = .fileName
                            Case 2
                                mskFileNameDettaglio = .fileName
                        End Select
'                End Select
                #End If
            End With
        Case fcImport
'            Select Case gmstrCodInstallazione
'                Case ciInst1
            #If VersionID = 1 Then
                    ImportRecords001
'                Case ciInst2
            #ElseIf (VersionID = 2) Then 'Or (VersionID = 5) Then
                    ImportRecords002
'            End Select
            #End If
            cmdImport.Enabled = False
        Case fcOpen
'            Select Case gmstrCodInstallazione
'                Case ciInst1
            #If VersionID = 1 Then
                    OpenTextFile001
'                Case ciInst2
            #ElseIf (VersionID = 2) Then 'Or (VersionID = 5) Then
                    OpenTextFile002
'            End Select
            #End If
        Case fcPrint
            If State = fsEdit Then
            End If
        Case fcExit
            Unload Me
    End Select
End Sub






Private Sub ImportRecords001()
    On Error GoTo ErrHandler
    
    Dim strIDClienteDefault As String
    Dim strIDImballoDefault As String
    Dim strIDPosDefault As String
    Dim strIDParam As String
    Dim strIDDestinazione As String
    Dim strIDTipoVendita As String
    Dim strIDEnteEmittente As String
    Dim strIDOrder As String
    Dim strIDOrderAGIP As String
    Dim strIDPrevOrderAGIP As String
    Dim strIDProdotto As String
    Dim strIDRiga  As String
    Dim strCodDest As String
    Dim strCodDestPrec As String
    Dim strIDDeposito As String
    Dim strIDZona As String
    Dim strNewLocalita As String
    Dim strErrorMessage As String
    Dim lngCountOK As Long
    Dim lngCountFail As Long
    Dim lngCountToGo As Long
    Dim lngCountExcluded As Long
    Dim lngNumRows As Long
    Dim lngStartRow As Long
    Dim Row As Long
    Dim Col As Integer
    Dim i As Integer
    Dim blnNewOrder As Boolean
    Dim blnLocHasChanged As Boolean
    Dim blnDeleteOrder As Boolean
    Dim blnOrdineDaAggiornare As Boolean
    Dim blnConcErr As Boolean
    Dim blnNotUpdatable As Boolean
    Dim intRowError As Errors
    Dim ErrCode As Errors
    Dim C As CError
    ReDim ListNewLoc(0 To 0) As String
    
    Me.MousePointer = vbHourglass
    Me.Enabled = False
    ErrCode = NoError
    
    ' --- Leggo i valori di default del sistema
    strIDParam = ID_PARAM
    With AXX.DS
        .Fields(AXX.DS.KeyFields(1)).FindValue = strIDParam
        If Not AXX.Find Then
            GoTo ErrHandler
        End If
        If .HowMany = 0 Then
            ErrCode = InvalidHeader
            GoTo ErrHandler
        End If
        strIDImballoDefault = Trim$(.Fields("1").value)
        strIDPosDefault = Trim$(.Fields("2").value)
        strIDClienteDefault = Trim$(.Fields("3").value)
    End With
    
    If (strIDImballoDefault = vbNullString) Or _
       (strIDPosDefault = vbNullString) Or _
       (strIDClienteDefault = vbNullString) Then
        ErrCode = MissingDefaults
        GoTo ErrHandler
    End If
    ' ---
    
    lngNumRows = XA.UpperBound(1)
    lngStartRow = XA.LowerBound(1)
    If lngNumRows = 0 Then
        ErrCode = NoImportData
        GoTo ErrHandler
    End If
    lngCountToGo = lblLog(5)
    
    If IsUserGrandesCuentas Then
        If Not OCT.InvalidaOrdini(ID_SOC, Format(Date, "dd/mm/yyyy")) Then
            GoTo ErrHandler
        End If
    End If
    ' --- Ciclo di acquisizione dati
    For Row = lngStartRow To lngNumRows
        intRowError = NoError
        
        ' --- Leggo una riga del file degli ordini
        For Col = 0 To 29 ' 22 '17
            mobjInFile(Col + 1).value = XA(Row, Col)
        Next Col
        ' ---
        
        strIDOrderAGIP = mobjInFile("12").value
        If strIDOrderAGIP = vbNullString Then
            intRowError = BadIDOrderSoc
            GoTo LogResult
        End If
        If strIDOrderAGIP <> strIDPrevOrderAGIP Then
            ' --- Nuovo ordine : Inizio con l'inserimento di una nuova testata
            blnNewOrder = True
            strIDPrevOrderAGIP = strIDOrderAGIP
            blnOrdineDaAggiornare = False
            ' --- Controllo che l'ordine non sia già stato inserito
            If OCT.CheckSocCode(strIDOrderAGIP) Then
                If mobjInFile("19").value = "TD" Then
                    blnOrdineDaAggiornare = True
                    strIDOrder = OCT.GetIDBySocCode(mobjInFile("12").value)
                Else
                    strIDOrder = vbNullString
                    intRowError = OrderAlreadyExisting
                    GoTo LogResult
                End If
            End If
            ' ---
        Else
            blnNewOrder = False
        End If
        
        
        intRowError = CheckOrderRow001(strIDDeposito, strIDEnteEmittente, _
                                       strIDProdotto, strIDTipoVendita, _
                                       strIDZona, strCodDest)
        
        If intRowError <> NoError Then
            GoTo LogResult
        End If
        
        If blnOrdineDaAggiornare Then
            If blnNewOrder Then
                intRowError = CheckDestinazione001(strIDDestinazione, strCodDest, _
                                                   strIDClienteDefault, strIDDeposito, _
                                                   strNewLocalita)
                If (intRowError <> NoError) And _
                   (intRowError <> DestHasChanged) Then
                    GoTo LogResult
                End If
                ' --- Aggiorno la testata dell'ordine
                OCT.DS.Fields("0").FindValue = strIDOrder
                OCT.Find
                If OCT.DS.Fields("23").value = "A" Then
                    blnNotUpdatable = False
                    RecordToOCTUpdate OCT, mobjInFile, strIDDeposito, strIDDestinazione, _
                                      strIDTipoVendita, strIDEnteEmittente, strIDClienteDefault
                    OCT.BeginSave blnConcErr
                    OCT.DS.DA.EndTrans ETCOMMIT
                    OCT.DS.EndSave
                    strIDRiga = 10
                    If Not OCT.DeleteRows(strIDOrder) Then
                        intRowError = OrderInsertError
                        GoTo LogResult
                    End If
                Else
                    blnNotUpdatable = True
                End If
            End If
        Else
            If blnNewOrder Then
                blnNotUpdatable = False
                intRowError = CheckDestinazione001(strIDDestinazione, strCodDest, _
                                                   strIDClienteDefault, strIDDeposito, _
                                                   strNewLocalita)
                If (intRowError <> NoError) And _
                   (intRowError <> DestHasChanged) Then
                    GoTo LogResult
                End If
                ' --- Inserisco la testata dell'ordine
                RecordToOCT OCT, mobjInFile, strIDDeposito, strIDDestinazione, _
                            strIDTipoVendita, strIDEnteEmittente, strIDClienteDefault
                If Not OCT.Insert Then
                    intRowError = OrderInsertError
                    GoTo LogResult
                End If
                strIDOrder = OCT.DS.Fields("0").NewValue
                strIDRiga = 10
                ' ---
            End If
            
            ' --- Inserisco una nuova riga ordine
            If strIDOrder = vbNullString Then
                intRowError = OrderAlreadyExisting
                GoTo LogResult
            End If
            
        End If
        
        If Not blnNotUpdatable Then
            RecordToOCR OCR, mobjInFile, strIDOrder, strIDRiga, _
                        strIDProdotto, strIDImballoDefault, strIDPosDefault
            strIDRiga = strIDRiga + 10
            If Not OCR.Insert Then
                intRowError = RowOrderInsertError
                GoTo LogResult
            End If
        End If
        ' ---
LogResult:
        ' --- Scrivo il risultato dell'importazione
        If (intRowError <> NoError) And _
           (intRowError <> ZonaEsclusaImport) And _
           (intRowError <> DestHasChanged) Then
            lngCountFail = lngCountFail + 1
        End If
        If blnNewOrder Then
            If (intRowError = NoError) Or _
               (intRowError = DestHasChanged) Then
                lngCountOK = lngCountOK + 1
            Else
                If intRowError = ZonaEsclusaImport Then
                    lngCountExcluded = lngCountExcluded + 1
                End If
            End If
        
            lngCountToGo = lngCountToGo - 1
        End If
        
        lblLog(2) = lngCountOK
        lblLog(4) = lngCountFail
        lblLog(5) = lngCountToGo
        lblLog(6) = lngCountExcluded
        
        Select Case intRowError
            Case NoError
                strErrorMessage = "OK - Num. Ordine : " & strIDOrder
            Case DestHasChanged
                blnLocHasChanged = True
                ListNewLoc(UBound(ListNewLoc)) = "Dest. modificata: " & strIDDestinazione & " " & _
                                                 strNewLocalita
                strErrorMessage = "OK - Num. Ordine : " & strIDOrder & _
                                  " - " & ListNewLoc(UBound(ListNewLoc))
                ReDim Preserve ListNewLoc(UBound(ListNewLoc) + 1) As String
            Case OrderAlreadyExisting
                strErrorMessage = "Ordine già inserito"
            Case NoTipoVendita
                strErrorMessage = "Codice Tipo Vendita non trovato"
            Case NoZona
                strErrorMessage = "Codice Zona non trovato"
            Case NoDeposito
                strErrorMessage = "Codice Deposito non trovato"
            Case NoEnteEmittente
                strErrorMessage = "Codice Ente Emittente non trovato"
            Case NoProdotto
                strErrorMessage = "Codice Prodotto non trovato"
            Case NoDestinazione
                strErrorMessage = "Codice Destinazione mancante"
            Case BadIDOrderSoc
                strErrorMessage = "Numero Ordine mancante"
            Case DestInsertError
                strErrorMessage = "Errore durante l'importazione della destinazione"
            Case DestSocInsertError
                strErrorMessage = "Errore durante l'importazione della destinazione"
            Case OrderInsertError
                strErrorMessage = "Errore durante l'inserimento dell'ordine "
            Case RowOrderInsertError
                strErrorMessage = "Errore durante l'inserimento della riga di ordine"
            Case ZonaEsclusaImport
                strErrorMessage = "Zona esclusa dall'importazione ordini"
        End Select
        XALog.value(Row, COL_RESULT_LOG) = strErrorMessage
        With tdbLog
            .Refresh
            .bookmark = Row
        End With
        TDBGrid1.bookmark = Row
        ' ---
        
        DoEvents
    Next Row
    ' ---
    
    If IsUserGrandesCuentas Then
        If Not OCT.AnnullaOrdiniTD(ID_SOC, Format(Date, "dd/mm/yyyy")) Then
            GoTo ErrHandler
        End If
    End If
    
    mstrUltimoProgressivo = lblField(0)
    GoTo EndProc
ErrHandler:
    Dim strMessage As String
    Dim strInfo As String
    Dim lngNativeCode As Long
    
    'OCT.DS.DA.EndTrans 1
    Select Case ErrCode
        Case InvalidHeader
            strMessage = "Impossibile trovare il parametro di sistema."
            strInfo = "Codice Parametro: " & strIDParam
        Case MissingDefaults
            strMessage = "Parametri di sistema mancanti."
            strInfo = "Default mancanti:" & vbNewLine
            If (strIDImballoDefault = vbNullString) Then
                strInfo = strInfo & "Imballo" & vbNewLine
            End If
            If (strIDPosDefault = vbNullString) Then
                strInfo = strInfo & "Posizione Doganale" & vbNewLine
            End If
            If (strIDClienteDefault = vbNullString) Then
                strInfo = strInfo & "Codice Cliente" & vbNewLine
            End If
        Case NoImportData
            strMessage = "Non ci sono dati da importare."
            strInfo = vbNullString
        Case NoError
            Select Case Err.Number
                Case Else
                    ErrCode = UnknownError
                    strMessage = "Errore Sconosciuto."
                    strInfo = Err.Description
                    lngNativeCode = Err.Number
            End Select
    End Select
    Set C = ErrH.SetError(ErrCode, "frmOCTImport", _
                "ImportRecords", lngNativeCode, strMessage, 10, strInfo)
    DisplayError ErrH.LastError
EndProc:
    If blnLocHasChanged Then
        If MsgBox("Alcune destinazioni sono state modificate. Vuoi stamparne l'elenco?", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
            For i = 0 To UBound(ListNewLoc) - 1
                Printer.Print
                Printer.Print
                Printer.Print ListNewLoc(i)
            Next i
            Printer.EndDoc
        End If
    End If
    Me.Enabled = True
    Me.MousePointer = vbDefault
End Sub

Private Sub ImportRecords004()
    On Error GoTo ErrHandler
    
    Dim strIDClienteDefault As String
    Dim strIDImballoDefault As String
    Dim strIDPosDefault As String
    Dim strIDParam As String
    Dim strIDDestinazione As String
    Dim strIDTipoVendita As String
    Dim strIDEnteEmittente As String
    Dim strIDOrder As String
    Dim strIDOrderAGIP As String
    Dim strIDPrevOrderAGIP As String
    Dim strIDProdotto As String
    Dim strIDRiga  As String
    Dim strCodDest As String
    Dim strCodDestPrec As String
    Dim strIDDeposito As String
    Dim strIDZona As String
    Dim strNewLocalita As String
    Dim strErrorMessage As String
    Dim lngCountOK As Long
    Dim lngCountFail As Long
    Dim lngCountToGo As Long
    Dim lngCountExcluded As Long
    Dim lngNumRows As Long
    Dim lngStartRow As Long
    Dim Row As Long
    Dim Col As Integer
    Dim i As Integer
    Dim blnNewOrder As Boolean
    Dim blnLocHasChanged As Boolean
    Dim intRowError As Errors
    Dim ErrCode As Errors
    Dim C As CError
    ReDim ListNewLoc(0 To 0) As String
    
    Me.MousePointer = vbHourglass
    Me.Enabled = False
    ErrCode = NoError
    
    ' --- Leggo i valori di default del sistema
    strIDParam = ID_PARAM
    With AXX.DS
        .Fields(AXX.DS.KeyFields(1)).FindValue = strIDParam
        If Not AXX.Find Then
            GoTo ErrHandler
        End If
        If .HowMany = 0 Then
            ErrCode = InvalidHeader
            GoTo ErrHandler
        End If
        strIDImballoDefault = Trim$(.Fields("1").value)
        strIDPosDefault = Trim$(.Fields("2").value)
        strIDClienteDefault = Trim$(.Fields("3").value)
    End With
    
    If (strIDImballoDefault = vbNullString) Or _
       (strIDPosDefault = vbNullString) Or _
       (strIDClienteDefault = vbNullString) Then
        ErrCode = MissingDefaults
        GoTo ErrHandler
    End If
    ' ---
    
    lngNumRows = XA.UpperBound(1)
    lngStartRow = XA.LowerBound(1)
    If lngNumRows = 0 Then
        ErrCode = NoImportData
        GoTo ErrHandler
    End If
    lngCountToGo = lblLog(5)
    
    ' --- Ciclo di acquisizione dati
    For Row = lngStartRow To lngNumRows
        intRowError = NoError
        
        ' --- Leggo una riga del file degli ordini
        For Col = 0 To 17
            mobjInFile(Col + 1).value = XA(Row, Col)
        Next Col
        ' ---
        
        strIDOrderAGIP = mobjInFile("12").value
        If strIDOrderAGIP = vbNullString Then
            intRowError = BadIDOrderSoc
            GoTo LogResult
        End If
        If strIDOrderAGIP <> strIDPrevOrderAGIP Then
            ' --- Nuovo ordine : Inizio con l'inserimento di una nuova testata
            blnNewOrder = True
            strIDPrevOrderAGIP = strIDOrderAGIP
            ' --- Controllo che l'ordine non sia già stato inserito
            If OCT.CheckSocCode(strIDOrderAGIP) Then
                strIDOrder = vbNullString
                intRowError = OrderAlreadyExisting
                GoTo LogResult
            End If
            ' ---
        Else
            blnNewOrder = False
        End If
        
        
        intRowError = CheckOrderRow001(strIDDeposito, strIDEnteEmittente, _
                                       strIDProdotto, strIDTipoVendita, _
                                       strIDZona, strCodDest)
        
        If intRowError <> NoError Then
            GoTo LogResult
        End If
        
        
        If blnNewOrder Then
            intRowError = CheckDestinazione001(strIDDestinazione, strCodDest, _
                                               strIDClienteDefault, strIDDeposito, _
                                               strNewLocalita)
            If (intRowError <> NoError) And _
               (intRowError <> DestHasChanged) Then
                GoTo LogResult
            End If
            ' --- Inserisco la testata dell'ordine
            RecordToOCT OCT, mobjInFile, strIDDeposito, strIDDestinazione, _
                        strIDTipoVendita, strIDEnteEmittente, strIDClienteDefault
            If Not OCT.Insert Then
                intRowError = OrderInsertError
                GoTo LogResult
            End If
            strIDOrder = OCT.DS.Fields("0").NewValue
            strIDRiga = 10
            ' ---
        End If
        
        ' --- Inserisco una nuova riga ordine
        If strIDOrder = vbNullString Then
            intRowError = OrderAlreadyExisting
            GoTo LogResult
        End If
        
        RecordToOCR OCR, mobjInFile, strIDOrder, strIDRiga, _
                    strIDProdotto, strIDImballoDefault, strIDPosDefault
        strIDRiga = strIDRiga + 10
        If Not OCR.Insert Then
            intRowError = RowOrderInsertError
            GoTo LogResult
        End If
        ' ---
        
LogResult:
        ' --- Scrivo il risultato dell'importazione
        If (intRowError <> NoError) And _
           (intRowError <> ZonaEsclusaImport) And _
           (intRowError <> DestHasChanged) Then
            lngCountFail = lngCountFail + 1
        End If
        If blnNewOrder Then
            If (intRowError = NoError) Or _
               (intRowError = DestHasChanged) Then
                lngCountOK = lngCountOK + 1
            Else
                If intRowError = ZonaEsclusaImport Then
                    lngCountExcluded = lngCountExcluded + 1
                End If
            End If
        
            lngCountToGo = lngCountToGo - 1
        End If
        
        lblLog(2) = lngCountOK
        lblLog(4) = lngCountFail
        lblLog(5) = lngCountToGo
        lblLog(6) = lngCountExcluded
        
        Select Case intRowError
            Case NoError
                strErrorMessage = "OK - Num. Ordine : " & strIDOrder
            Case DestHasChanged
                blnLocHasChanged = True
                ListNewLoc(UBound(ListNewLoc)) = "Dest. modificata: " & strIDDestinazione & " " & _
                                                 strNewLocalita
                strErrorMessage = "OK - Num. Ordine : " & strIDOrder & _
                                  " - " & ListNewLoc(UBound(ListNewLoc))
                ReDim Preserve ListNewLoc(UBound(ListNewLoc) + 1) As String
            Case OrderAlreadyExisting
                strErrorMessage = "Ordine già inserito"
            Case NoTipoVendita
                strErrorMessage = "Codice Tipo Vendita non trovato"
            Case NoZona
                strErrorMessage = "Codice Zona non trovato"
            Case NoDeposito
                strErrorMessage = "Codice Deposito non trovato"
            Case NoEnteEmittente
                strErrorMessage = "Codice Ente Emittente non trovato"
            Case NoProdotto
                strErrorMessage = "Codice Prodotto non trovato"
            Case NoDestinazione
                strErrorMessage = "Codice Destinazione mancante"
            Case BadIDOrderSoc
                strErrorMessage = "Numero Ordine mancante"
            Case DestInsertError
                strErrorMessage = "Errore durante l'importazione della destinazione"
            Case DestSocInsertError
                strErrorMessage = "Errore durante l'importazione della destinazione"
            Case OrderInsertError
                strErrorMessage = "Errore durante l'inserimento dell'ordine "
            Case RowOrderInsertError
                strErrorMessage = "Errore durante l'inserimento della riga di ordine"
            Case ZonaEsclusaImport
                strErrorMessage = "Zona esclusa dall'importazione ordini"
        End Select
        XALog.value(Row, COL_RESULT_LOG) = strErrorMessage
        With tdbLog
            .Refresh
            .bookmark = Row
        End With
        TDBGrid1.bookmark = Row
        ' ---
        
        DoEvents
    Next Row
    ' ---
    
    GoTo EndProc
ErrHandler:
    Dim strMessage As String
    Dim strInfo As String
    Dim lngNativeCode As Long
    
    'OCT.DS.DA.EndTrans 1
    Select Case ErrCode
        Case InvalidHeader
            strMessage = "Impossibile trovare il parametro di sistema."
            strInfo = "Codice Parametro: " & strIDParam
        Case MissingDefaults
            strMessage = "Parametri di sistema mancanti."
            strInfo = "Default mancanti:" & vbNewLine
            If (strIDImballoDefault = vbNullString) Then
                strInfo = strInfo & "Imballo" & vbNewLine
            End If
            If (strIDPosDefault = vbNullString) Then
                strInfo = strInfo & "Posizione Doganale" & vbNewLine
            End If
            If (strIDClienteDefault = vbNullString) Then
                strInfo = strInfo & "Codice Cliente" & vbNewLine
            End If
        Case NoImportData
            strMessage = "Non ci sono dati da importare."
            strInfo = vbNullString
        Case NoError
            Select Case Err.Number
                Case Else
                    ErrCode = UnknownError
                    strMessage = "Errore Sconosciuto."
                    strInfo = Err.Description
                    lngNativeCode = Err.Number
            End Select
    End Select
    Set C = ErrH.SetError(ErrCode, "frmOCTImport", _
                "ImportRecords", lngNativeCode, strMessage, 10, strInfo)
    DisplayError ErrH.LastError
EndProc:
    If blnLocHasChanged Then
        If MsgBox("Alcune destinazioni sono state modificate. Vuoi stamparne l'elenco?", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
            For i = 0 To UBound(ListNewLoc) - 1
                Printer.Print
                Printer.Print
                Printer.Print ListNewLoc(i)
            Next i
            Printer.EndDoc
        End If
    End If
    Me.Enabled = True
    Me.MousePointer = vbDefault
End Sub


Private Sub ImportRecords002()
    On Error GoTo ErrHandler
    
    Dim strIDClienteDefault As String
    Dim strIDImballoDefault As String
    Dim strIDPosDoganale As String
    Dim strIDParam As String
    Dim strIDDestinazione As String
    Dim strIDTipoVendita As String
    Dim strIDEnteEmittente As String
    Dim strIDOrder As String
    Dim strIDOrderSoc As String
    Dim strIDPrevOrderAGIP As String
    Dim strIDProdotto As String
    Dim strIDRiga  As String
    Dim strCodDest As String
    Dim strCodDestPrec As String
    Dim strIDDeposito As String
    Dim strIDZona As String
    Dim strErrorMessage As String
    Dim lngCountOK As Long
    Dim lngCountFail As Long
    Dim lngCountToGo As Long
    Dim lngCountExcluded As Long
    Dim lngNumRows As Long
    Dim lngNumOrderRows As Long
    Dim lngStartRow As Long
    Dim dblQtaProdotto As Double
    Dim lngCurrentOrderRow As Long
    Dim Row As Long
    Dim Col As Integer
    Dim blnNewOrder As Boolean
    Dim blnRowInsertFlag As Boolean
    Dim intRowError As Errors
    Dim ErrCode As Errors
    Dim C As CError
    
    Me.MousePointer = vbHourglass
    Me.Enabled = False
    ErrCode = NoError
    
    ' --- Leggo i valori di default del sistema
    strIDParam = ID_PARAM
    With AXX.DS
        .Fields(AXX.DS.KeyFields(1)).FindValue = strIDParam
        If Not AXX.Find Then
            GoTo ErrHandler
        End If
        If .HowMany = 0 Then
            ErrCode = InvalidHeader
            GoTo ErrHandler
        End If
        strIDImballoDefault = Trim$(.Fields("1").value)
        strIDClienteDefault = Trim$(.Fields("3").value)
    End With
    
    If (strIDImballoDefault = vbNullString) Or _
       (strIDClienteDefault = vbNullString) Then
        ErrCode = MissingDefaults
        GoTo ErrHandler
    End If
    ' ---
    
    lngNumRows = XALog.UpperBound(1)
    lngStartRow = XALog.LowerBound(1)
    lngCurrentOrderRow = XA.LowerBound(1)
    lngNumOrderRows = XA.UpperBound(1)
    If lngNumRows = 0 Then
        ErrCode = NoImportData
        GoTo ErrHandler
    End If
    lngCountToGo = lblLog(5)
    intRowError = NoError
    
    ' --- Ciclo di acquisizione dati
    For Row = lngStartRow To lngNumRows
        blnNewOrder = False
        
        ' --- Leggo una riga di dettaglio
        For Col = 0 To 9
            mobjInFileDettaglio(Col + 2).value = XALog.value(Row, Col)
        Next Col
        ' ---
        
        ' --- Il codice ordine è cambiato ?
        If strIDOrderSoc <> mobjInFileDettaglio("3").value Then
            intRowError = NoError
            
            If blnRowInsertFlag Then
                OCT.DS.DA.EndTrans 0
            Else
                OCT.DS.DA.EndTrans 1
            End If
            
            lngCurrentOrderRow = 1
            Do While XA(lngCurrentOrderRow, 1) <> mobjInFileDettaglio("3").value
                lngCurrentOrderRow = lngCurrentOrderRow + 1
            Loop
            For Col = 0 To 15
                mobjInFile(Col + 2).value = XA(lngCurrentOrderRow, Col)
            Next Col
            
            lngCurrentOrderRow = lngCurrentOrderRow + 1
            blnNewOrder = True
            blnRowInsertFlag = False
            
            ' --- Controllo che esistano Deposito, Zona, Ente Emittente, Prodotto
            strIDEnteEmittente = AEE.DecodeFirmCode(ID_SOC, mobjInFile("7").value, True)
    
            strIDZona = mobjInFile("7").value
            AZO.DS.Fields("Codice").FindValue = strIDZona
            If Not AZO.Find Then
                intRowError = NoZona
                GoTo LogResult
            End If
            If AZO.DS.HowMany = 0 Then
                intRowError = NoZona
                GoTo LogResult
            End If
            
            ' --- Controllo che la zona non sia da escludere
            ' --- dall'importazione
            If AZO.DS.Fields("2").value <> 0 Then
                intRowError = ZonaEsclusaImport
                GoTo LogResult
            End If
            ' ---
            
            If strIDEnteEmittente = vbNullString Then
                intRowError = NoEnteEmittente
                GoTo LogResult
            End If
            
            strCodDest = mobjInFile("5").value & mobjInFile("6").value
            If strCodDest = vbNullString Then
                intRowError = NoDestinazione
                GoTo LogResult
            End If
            
            strIDOrderSoc = mobjInFile("3").value
            If strIDOrderSoc = vbNullString Then
                intRowError = BadIDOrderSoc
                GoTo LogResult
            End If
        
            ' --- Controllo che l'ordine non sia già stato inserito
            If OCT.CheckSocCode(strIDOrderSoc) Then
                intRowError = OrderAlreadyExisting
                GoTo LogResult
            End If
            ' ---
            
            OCT.DS.DA.BeginTransaction
            
            ' --- Cerco la destinazione
            strIDDestinazione = ADZ.DecodeFirmCode(ID_SOC, strCodDest, True)
            
            If strIDDestinazione = vbNullString Then
                intRowError = NoDestinazione
                GoTo LogResult
            End If
            
            ' --- Inserisco la testata dell'ordine
            RecordToOCT002 OCT, mobjInFile, strIDDestinazione, _
                           strIDTipoVendita, strIDEnteEmittente, strIDClienteDefault
            If Not OCT.Insert Then
                intRowError = OrderInsertError
                GoTo LogResult
            End If
            strIDOrder = OCT.DS.Fields("0").NewValue
            strIDRiga = 10
            blnRowInsertFlag = False
            ' ---
        End If
        
        ' --- Inserisco il dettaglio dell'ordine
        If (strIDOrder <> vbNullString) And (intRowError = NoError) Then
            
            strIDProdotto = APR.DecodeFirmCode(ID_SOC, mobjInFileDettaglio("6").value, True)
            If strIDProdotto = vbNullString Then
                intRowError = NoProdotto
                GoTo LogResult
            End If
            
            strIDPosDoganale = APD.DecodeFirmCode(ID_SOC, mobjInFileDettaglio("8").value, True)
            If strIDPosDoganale = vbNullString Then
                intRowError = NoPosizioneDoganale
                GoTo LogResult
            End If
            
            If Len(Trim(mobjInFileDettaglio("11").value)) > 0 Then
                dblQtaProdotto = mobjInFileDettaglio("11").value
            Else
                dblQtaProdotto = mobjInFileDettaglio("10").value
            End If
            
            If dblQtaProdotto <> 0 Then
                dblQtaProdotto = dblQtaProdotto * 1000
                RecordToOCR002 OCR, mobjInFileDettaglio, strIDOrder, strIDRiga, _
                               strIDProdotto, strIDImballoDefault, strIDPosDoganale, _
                               dblQtaProdotto
                strIDRiga = strIDRiga + 10
                If Not OCR.Insert Then
                    intRowError = RowOrderInsertError
                    GoTo LogResult
                End If
                blnRowInsertFlag = True
            End If
        End If
        ' ---
        
LogResult:
        ' --- Scrivo il risultato dell'importazione
        If blnNewOrder Then
            If intRowError = NoError Then
                OCT.DS.DA.EndTrans 0
            Else
                OCT.DS.DA.EndTrans 1
            End If
            
            If intRowError = NoError Then
                lngCountOK = lngCountOK + 1
            Else
                If (intRowError = ZonaEsclusaImport) Then
                    lngCountExcluded = lngCountExcluded + 1
                Else
                    lngCountFail = lngCountFail + 1
                End If
            End If
            lngCountToGo = lngCountToGo - 1
        End If
        lblLog(2) = lngCountOK
        lblLog(4) = lngCountFail
        lblLog(5) = lngCountToGo
        lblLog(6) = lngCountExcluded
        
        Select Case intRowError
            Case NoError
                strErrorMessage = "OK - Num. Ordine : " & strIDOrder
            Case NoTipoVendita
                strErrorMessage = "Codice Tipo Vendita non trovato"
            Case NoZona
                strErrorMessage = "Codice Zona non trovato"
            Case NoDeposito
                strErrorMessage = "Codice Deposito non trovato"
            Case NoEnteEmittente
                strErrorMessage = "Codice Ente Emittente non trovato"
            Case NoProdotto
                strErrorMessage = "Codice Prodotto non trovato"
            Case NoPosizioneDoganale
                strErrorMessage = "Codice Posizione Doganale non trovato"
            Case NoDestinazione
                strErrorMessage = "Codice Destinazione mancante"
            Case BadIDOrderSoc
                strErrorMessage = "Numero Ordine mancante"
            Case DestInsertError
                strErrorMessage = "Errore durante l'importazione della destinazione"
            Case DestSocInsertError
                strErrorMessage = "Errore durante l'importazione della destinazione"
            Case OrderInsertError
                strErrorMessage = "Errore durante l'inserimento dell'ordine "
            Case RowOrderInsertError
                strErrorMessage = "Errore durante l'inserimento della riga di ordine"
            Case ZonaEsclusaImport
                strErrorMessage = "Zona esclusa dall'importazione ordini"
            Case OrderAlreadyExisting
                strErrorMessage = "Ordine già inserito"
        End Select
        XALog.value(Row, 10) = strErrorMessage
        With tdbLog
            .Refresh
            .bookmark = Row
        End With
        If lngCurrentOrderRow <= lngNumOrderRows Then
            TDBGrid2.bookmark = lngCurrentOrderRow
        End If
        ' ---
        
        DoEvents
    Next Row
    ' ---
    If blnRowInsertFlag Then
        OCT.DS.DA.EndTrans 0
    Else
        OCT.DS.DA.EndTrans 1
    End If
    
    GoTo EndProc
ErrHandler:
    Dim strMessage As String
    Dim strInfo As String
    Dim lngNativeCode As Long
    
    OCT.DS.DA.EndTrans 1
    Select Case ErrCode
        Case InvalidHeader
            strMessage = "Impossibile trovare il parametro di sistema."
            strInfo = "Codice Parametro: " & strIDParam
        Case MissingDefaults
            strMessage = "Parametri di sistema mancanti."
            strInfo = "Default mancanti:" & vbNewLine
            If (strIDImballoDefault = vbNullString) Then
                strInfo = strInfo & "Imballo" & vbNewLine
            End If
            If (strIDClienteDefault = vbNullString) Then
                strInfo = strInfo & "Codice Cliente" & vbNewLine
            End If
        Case NoImportData
            strMessage = "Non ci sono dati da importare."
            strInfo = vbNullString
        Case NoError
            Select Case Err.Number
                Case Else
                    ErrCode = UnknownError
                    strMessage = "Errore Sconosciuto."
                    strInfo = Err.Description
                    lngNativeCode = Err.Number
            End Select
    End Select
    Set C = ErrH.SetError(ErrCode, "frmOCTImport", _
                "ImportRecords", lngNativeCode, strMessage, 10, strInfo)
    DisplayError ErrH.LastError
EndProc:
    OCT.DS.DA.EndTrans 0
    Me.Enabled = True
    Me.MousePointer = vbDefault
End Sub


Private Sub OpenTextFile001()
    On Error GoTo ErrHandler
    
    Dim strFileName As String
    Dim strLine As String
    Dim strEnte As String
    Dim strPreviousOrder As String
    Dim varBuffer As Variant
    Dim blnValidHeader As Boolean
    Dim blnIDOk As Boolean
    Dim ErrCode As Errors
    Dim C As CError
    Dim i As Integer
    Dim lngRows As Long
    Dim lngCountOrders As Long
    
    Me.MousePointer = vbHourglass
    ErrCode = NoError
    strFileName = mskFileName
    
    If Not mobjInFile.OpenFile(strFileName, 0) Then
        ErrCode = FileNotFound
        GoTo ErrHandler
    End If
    For i = 0 To 3
        lblField(i) = vbNullString
    Next i
    cmdImport.Enabled = False
    Set XA = New XArrayDB
    Set XALog = New XArrayDB
    lngRows = 1
    Do While Not mobjInFile.EOF
        varBuffer = mobjInFile.GetLine
        If varBuffer(0) = "1" Then
            blnValidHeader = True
            mstrIDEnte = CStr(varBuffer(1))
            strEnte = AEE.DecodeFirmCode(ID_SOC, mstrIDEnte)
            If strEnte <> vbNullString Then
                lblField(1) = strEnte
                blnIDOk = True
            End If
            lblField(0) = varBuffer(2)
            lblField(2) = varBuffer(3)
            lblField(3) = varBuffer(4)
            If IsNumeric(mstrUltimoProgressivo) Then
                If lblField(0) <> CStr(CLng(mstrUltimoProgressivo) + 1) Then
                    MsgBox "Il numero del file atteso è:  " & (CLng(mstrUltimoProgressivo) + 1), vbOKOnly + vbExclamation, "Attenzione"
                End If
            End If
        Else
            With XA
                .ReDim 1, lngRows, 0, 29 ' 22
                For i = 0 To 29 ' 22
                    .value(.UpperBound(1), i) = varBuffer(i + 1)
                Next i
            End With
            With XALog
                .ReDim 1, lngRows, 0, 10 '2
                .value(.UpperBound(1), COL_COD_ORD_LOG) = varBuffer(12) '11)
                .value(.UpperBound(1), COL_COD_PROD_LOG) = varBuffer(17) '16)
            End With
'            If varBuffer(11) <> strPreviousOrder Then
            If varBuffer(12) <> strPreviousOrder Then
'                strPreviousOrder = varBuffer(11)
                strPreviousOrder = varBuffer(12)
                lngCountOrders = lngCountOrders + 1
            End If
            lngRows = lngRows + 1
        End If
    Loop

    If Not blnValidHeader Then
        Set C = ErrH.SetError(InvalidHeader, "frmImportOrders", _
                    "OpenTextFile", 0, "Il file ha un formato non valido.", 10, "Nome del file: " & strFileName)
        DisplayError ErrH.LastError

    Else
        If Not blnIDOk Then
            Set C = ErrH.SetError(InvalidCode, "frmImportOrders", _
                    "OpenTextFile", 0, "Codice Ente Emittente non trovato.", 10, "Codice Ente: " & CStr(varBuffer(1)))
            DisplayError ErrH.LastError
        Else
            cmdImport.Enabled = True
            XA.QuickSort 1, XA.UpperBound(1), _
                         11, XORDER_ASCEND, XTYPE_STRING
            With TDBGrid1
                .Array = XA
                .ReBind
                .bookmark = 1
            End With
            XALog.QuickSort 1, XALog.UpperBound(1), _
                            1, XORDER_ASCEND, XTYPE_STRING
            With tdbLog
                .Array = XALog
                .ReBind
                .bookmark = 1
            End With
            lblLog(2) = 0
            lblLog(4) = 0
            lblLog(5) = lngCountOrders
            lblLog(6) = 0
        End If
    End If
    GoTo EndProc
ErrHandler:
    Select Case ErrCode
        Case FileNotFound
            Set C = ErrH.SetError(FileNotFound, "frmOCTImport", _
                    "OpenTextFile", Err.Number, "Il file non è stato trovato.", 10, "Nome del file: " & strFileName)
        Case NoError
            Select Case Err.Number
                Case Else
                    Set C = ErrH.SetError(UnknownError, "frmOCTImport", _
                            "OpenTextFile", Err.Number, "Errore sconosciuto", 10, vbNullString)
            End Select
    End Select
    C.Descriptions.Add Err.Description
    DisplayError ErrH.LastError
EndProc:
    mobjInFile.CloseFile
    Me.MousePointer = vbDefault
End Sub



Private Sub OpenTextFile002()
    On Error GoTo ErrHandler
    
    Dim strFileName As String
    Dim strLine As String
    Dim strEnte As String
    Dim strIDOrdine As String
    Dim strIDEnteDefault As String
    Dim varBuffer As Variant
    Dim blnValidHeader As Boolean
    Dim blnIDOk As Boolean
    Dim blnFound As Boolean
    Dim ErrCode As Errors
    Dim C As CError
    Dim i As Integer
    Dim j As Long
    Dim lngRows As Long
    Dim lngCountOrders As Long
    Dim lngUbound As Long
    
    Me.MousePointer = vbHourglass
    ErrCode = NoError
    strFileName = mskFileName
    
    If Not mobjInFile.OpenFile(strFileName, 0) Then
        ErrCode = FileNotFound
        GoTo ErrHandler
    End If
    For i = 0 To 3
        lblField(i) = vbNullString
    Next i
    cmdImport.Enabled = False
    
    ' --- Leggo i valori di default del sistema
    With AXX.DS
        .Fields(AXX.DS.KeyFields(1)).FindValue = ID_PARAM
        If Not AXX.Find Then
            GoTo ErrHandler
        End If
        If .HowMany = 0 Then
            ErrCode = InvalidHeader
            GoTo ErrHandler
        End If
        strIDEnteDefault = Trim$(.Fields("16").value)
        strEnte = AEE.DecodeFirmCode(ID_SOC, strIDEnteDefault)
        If strEnte <> vbNullString Then
            lblField(1) = strEnte
        End If
    End With
    
    If (strIDEnteDefault = vbNullString) Then
        ErrCode = MissingDefaults
        GoTo ErrHandler
    End If
    ' ---

    Set XA = New XArrayDB
    Set XALog = New XArrayDB
    lngRows = 1
    lngCountOrders = 1
    Do While Not mobjInFile.EOF
        mobjInFile.ParseFixedLenFields mobjInFile.GetLine(False), 1, 17
        If mobjInFile("7").value = strIDEnteDefault Then
            With XA
                .ReDim 1, lngCountOrders, 0, 15
                For i = 0 To 15
                    .value(.UpperBound(1), i) = mobjInFile(i + 2).value 'varBuffer(i + 1)
                Next i
            End With
            lngCountOrders = lngCountOrders + 1
        End If
        DoEvents
    Loop
    
    If lngCountOrders = 1 Then
        ErrCode = NoImportData
        GoTo ErrHandler
    End If
    strFileName = mskFileNameDettaglio
    If Not mobjInFileDettaglio.OpenFile(strFileName, 0) Then
        ErrCode = FileNotFound
        GoTo ErrHandler
    End If
    XALog.ReDim 1, 1, 0, 10
    Do While Not mobjInFileDettaglio.EOF
        mobjInFileDettaglio.ParseFixedLenFields mobjInFileDettaglio.GetLine(False), 1, 12
        strIDOrdine = mobjInFileDettaglio("3").value
        blnFound = False
        For i = 1 To XA.UpperBound(1)
            If strIDOrdine = XA.value(i, 1) Then
                blnFound = True
                Exit For
            End If
        Next i
        If blnFound Then
            With XALog
                j = 1
                lngUbound = .UpperBound(1)
                Do While (strIDOrdine >= .value(j, 1)) And (j < lngUbound)
                    j = j + 1
                Loop
                If lngRows <> 1 Then
                    If j = lngUbound Then
                        j = j + 1
                    End If
                    .Insert 1, j
                Else
                    j = 1
                End If
                
                For i = 0 To 9
                    .value(j, i) = mobjInFileDettaglio(i + 2).value  'varBuffer(i + 1)
                Next i
            End With
            lngRows = lngRows + 1
        End If
        DoEvents
    Loop
    
    cmdImport.Enabled = True
    With TDBGrid2
        .Array = XA
        .ReBind
        .bookmark = 1
    End With
    With tdbLog
        .Array = XALog
        .ReBind
        .bookmark = 1
    End With
    lblLog(2) = 0
    lblLog(4) = 0
    lblLog(5) = lngCountOrders - 1
    lblLog(6) = 0
    GoTo EndProc
ErrHandler:
    Select Case ErrCode
        Case FileNotFound
            Set C = ErrH.SetError(FileNotFound, "frmOCTImport", _
                    "OpenTextFile", Err.Number, "Il file non è stato trovato.", 10, "Nome del file: " & strFileName)
        Case MissingDefaults
            Set C = ErrH.SetError(MissingDefaults, "frmOCTImport", _
                    "OpenTextFile", Err.Number, "Parametri di sistema mancanti", 10, "Nome del file: " & strFileName)
        Case NoImportData
            Set C = ErrH.SetError(NoImportData, "frmOCTImport", _
                    "OpenTextFile", Err.Number, "Nessun ordine da importare.", 10, "Nome del file: " & strFileName)
        Case NoError
            Select Case Err.Number
                Case Else
                    Set C = ErrH.SetError(UnknownError, "frmOCTImport", _
                            "OpenTextFile", Err.Number, "Errore sconosciuto", 10, vbNullString)
            End Select
    End Select
    C.Descriptions.Add Err.Description
    DisplayError ErrH.LastError
EndProc:
    mobjInFile.CloseFile
    Me.MousePointer = vbDefault
End Sub


Private Sub RecordToADZ(ADZ As CDestinazione, InFile As CTextFile, IDClienteDefault As String, IDDeposito As String)
    Dim strIDDestinazione As String
    
    With ADZ.DS.Fields
        .ResetFields
        strIDDestinazione = Format$(InFile("4").value, "00000000")
        .Item("0").NewValue = strIDDestinazione 'Trim$(Left$(InFile("4").value, .Item("0").MaxLength))
        .Item("0").Updated = True
        .Item("1").NewValue = Trim$(Left$(InFile("5").value, .Item("1").MaxLength))
        .Item("1").Updated = True
        .Item("3").NewValue = Trim$(Left$(InFile("5").value, .Item("3").MaxLength))
        .Item("3").Updated = True
        .Item("4").NewValue = Trim$(Left$(InFile("6").value, .Item("4").MaxLength))
        .Item("4").Updated = True
        .Item("5").NewValue = Trim$(Left$(InFile("7").value, .Item("5").MaxLength))
        .Item("5").Updated = True
        .Item("7").NewValue = Trim$(Left$(InFile("8").value, .Item("7").MaxLength))
        .Item("7").Updated = True
        .Item("8").NewValue = IDClienteDefault
        .Item("8").Updated = True
        .Item("15").NewValue = IDDeposito
        .Item("15").Updated = True
        .Item("16").NewValue = Trim$(Left$(InFile("2").value, .Item("16").MaxLength))
        .Item("16").Updated = True
        .Item("18").NewValue = Trim$(Left$(InFile("9").value, .Item("18").MaxLength))
        .Item("18").Updated = True
        ' Km Raggio
        .Item("29").NewValue = InFile("15").value
        .Item("29").Updated = True
        
    End With
    With ADZ.DS.MultiSoc.DS.Fields
        .ResetFields
        .Item("Codice").NewValue = strIDDestinazione 'Trim$(Left$(InFile("4").value, .Item("Codice").MaxLength))
        .Item("Codice").Updated = True
        .Item("Società").NewValue = ID_SOC
        .Item("Società").Updated = True
        .Item("CodPerSoc").NewValue = Trim$(Left$(InFile("4").value, .Item("CodPerSoc").MaxLength))
        .Item("CodPerSoc").Updated = True
    End With
End Sub

Private Sub RecordToOCR(OCR As COrdineRiga, InFile As CTextFile, IDOrdine As String, _
                        IDRiga As String, IDProdotto As String, IDImballo As String, _
                        IDPosDoganale As String)
    Dim dblQta As Double
    
    dblQta = CDbl(Trim$(InFile("18").value))
    With OCR.DS.Fields
        .ResetFields
        .Item("0").NewValue = IDOrdine
        .Item("0").Updated = True
        .Item("1").NewValue = IDRiga
        .Item("1").Updated = True
        .Item("2").NewValue = IDProdotto
        .Item("2").Updated = True
        .Item("3").NewValue = IDImballo
        .Item("3").Updated = True
        .Item("4").NewValue = IDPosDoganale
        .Item("4").Updated = True
        .Item("5").NewValue = Trim$(Left$(InFile("14").value, .Item("5").MaxLength))
        .Item("5").Updated = True
        .Item("7").NewValue = dblQta
        .Item("7").Updated = True
        .Item("11").NewValue = dblQta
        .Item("11").Updated = True
        .Item("21").NewValue = ""
        .Item("21").Updated = True
        .Item("24").NewValue = Trim$(Left$(InFile("12").value, .Item("24").MaxLength))
        .Item("24").Updated = True
    End With
End Sub
Private Sub RecordToOCR002(OCR As COrdineRiga, InFile As CTextFile, IDOrdine As String, _
                        IDRiga As String, IDProdotto As String, IDImballo As String, _
                        IDPosDoganale As String, Qta As Double)
    
    With OCR.DS.Fields
        .ResetFields
        .Item("0").NewValue = IDOrdine
        .Item("0").Updated = True
        .Item("1").NewValue = IDRiga
        .Item("1").Updated = True
        .Item("2").NewValue = IDProdotto
        .Item("2").Updated = True
        .Item("3").NewValue = IDImballo
        .Item("3").Updated = True
        .Item("4").NewValue = IDPosDoganale
        .Item("4").Updated = True
        .Item("7").NewValue = Qta
        .Item("7").Updated = True
        .Item("11").NewValue = Qta
        .Item("11").Updated = True
        .Item("21").NewValue = ""
        .Item("21").Updated = True
        .Item("24").NewValue = Trim$(Left$(InFile("3").value, .Item("24").MaxLength))
        .Item("24").Updated = True
    End With
End Sub
Private Sub RecordToOCT(OCT As COrdine, InFile As CTextFile, _
                        IDDeposito As String, IDDestinazione As String, _
                        IDTipoVendita As String, IDEnteEmittente As String, _
                        IDCliente As String)
    With OCT.DS.Fields
        .ResetFields
        .Item("1").NewValue = ID_SOC
        .Item("1").Updated = True
        .Item("2").NewValue = IDEnteEmittente
        .Item("2").Updated = True
        .Item("3").NewValue = Format$(Date, "dd/mm/yyyy")
        .Item("3").Updated = True
        .Item("4").NewValue = Format$(Time, "hh:mm")
        .Item("4").Updated = True
        .Item("7").NewValue = Format$(InFile("16").value, "dd/mm/yyyy")
        .Item("7").Updated = True
        .Item("8").NewValue = Format$(InFile("16").value, "dd/mm/yyyy")
        .Item("8").Updated = True
        .Item("10").NewValue = Trim$(Left$(InFile("12").value, .Item("10").MaxLength))
        .Item("10").Updated = True
        .Item("12").NewValue = IDCliente
        .Item("12").Updated = True
        .Item("13").NewValue = IDDestinazione
        .Item("13").Updated = True
        .Item("14").NewValue = Trim$(Left$(InFile("10").value, .Item("14").MaxLength))
        .Item("14").Updated = True
        .Item("15").NewValue = Trim$(Left$(InFile("11").value, .Item("15").MaxLength))
        .Item("15").Updated = True
        .Item("16").NewValue = IDTipoVendita
        .Item("16").Updated = True
        
        ' Identifichiamo gli ordini da evadere come "Grandes Cuentas"
        If UCase(IDTipoVendita) = "TD" Then
            .Item("20").NewValue = "GC"
        Else
            .Item("20").NewValue = "RP"
        End If
        .Item("20").Updated = True
        
        .Item("17").NewValue = IDDeposito
        .Item("17").Updated = True
        .Item("18").NewValue = Trim$(Left$(InFile("2").value, .Item("18").MaxLength))
        .Item("18").Updated = True
        .Item("21").NewValue = "E"
'        .Item("36").NewValue = Trim$(Left$(InFile("24").value, .Item("36").MaxLength))
'        .Item("36").Updated = True
        .Item("21").Updated = True
        ' Nota 3
        .Item("36").NewValue = Trim$(Left$(InFile("24").value, .Item("36").MaxLength))
        .Item("36").Updated = True
        ' Urgenza
        .Item("25").NewValue = UCase(Left$(InFile("22").value, .Item("25").MaxLength))
        .Item("25").Updated = True
        ' Programmato
        .Item("28").NewValue = UCase(Left$(InFile("21").value, .Item("28").MaxLength))
        .Item("28").Updated = True
        ' Momento
        .Item("27").NewValue = UCase(Left$(InFile("23").value, .Item("27").MaxLength))
        .Item("27").Updated = True
        ' Km Raggio
        If Len(InFile("15").value) > 0 Then
            .Item("37").NewValue = InFile("15").value
            .Item("37").Updated = True
        End If
        ' Data Grado
        .Item("52").NewValue = InFile("25").value
        .Item("52").Updated = True
    End With

End Sub

Private Sub RecordToOCTUpdate(OCT As COrdine, InFile As CTextFile, _
                              IDDeposito As String, IDDestinazione As String, _
                              IDTipoVendita As String, IDEnteEmittente As String, _
                              IDCliente As String)
    With OCT.DS.Fields
        .Item("7").NewValue = Format$(InFile("16").value, "dd/mm/yyyy")
        .Item("7").Updated = True
        .Item("8").NewValue = Format$(InFile("16").value, "dd/mm/yyyy")
        .Item("8").Updated = True
        .Item("12").NewValue = IDCliente
        .Item("12").Updated = True
        .Item("13").NewValue = IDDestinazione
        .Item("13").Updated = True
        .Item("14").NewValue = Trim$(Left$(InFile("10").value, .Item("14").MaxLength))
        .Item("14").Updated = True
        .Item("15").NewValue = Trim$(Left$(InFile("11").value, .Item("15").MaxLength))
        .Item("15").Updated = True
        .Item("16").NewValue = IDTipoVendita
        .Item("16").Updated = True
        ' Identifichiamo gli ordini da evadere come "Grandes Cuentas"
        If UCase(IDTipoVendita) = "TD" Then
            .Item("20").NewValue = "GC"
        Else
            .Item("20").NewValue = "RP"
        End If
        .Item("20").Updated = True
        
        .Item("17").NewValue = IDDeposito
        .Item("17").Updated = True
        .Item("18").NewValue = Trim$(Left$(InFile("2").value, .Item("18").MaxLength))
        .Item("18").Updated = True
        .Item("21").NewValue = "E"
'        .Item("36").NewValue = Trim$(Left$(InFile("24").value, .Item("36").MaxLength))
'        .Item("36").Updated = True
        .Item("21").Updated = True
        ' Nota 3
        .Item("36").NewValue = Trim$(Left$(InFile("24").value, .Item("36").MaxLength))
        .Item("36").Updated = True
        ' Urgenza
        .Item("25").NewValue = UCase(Left$(InFile("22").value, .Item("25").MaxLength))
        .Item("25").Updated = True
        ' Programmato
        .Item("28").NewValue = UCase(Left$(InFile("21").value, .Item("28").MaxLength))
        .Item("28").Updated = True
        ' Momento
        .Item("27").NewValue = UCase(Left$(InFile("23").value, .Item("27").MaxLength))
        .Item("27").Updated = True
        ' Raggruppamento Utente 4
        .Item("29").NewValue = ""
        .Item("29").Updated = True
        ' Km Raggio
        If Len(InFile("15").value) > 0 Then
            .Item("37").NewValue = InFile("15").value
            .Item("37").Updated = True
        End If
        ' Data Grado
        .Item("52").NewValue = InFile("25").value
        .Item("52").Updated = True
    End With

End Sub


Private Sub RecordToOCT002(OCT As COrdine, InFile As CTextFile, _
                            IDDestinazione As String, _
                            IDTipoVendita As String, IDEnteEmittente As String, _
                            IDCliente As String)
    Dim strDtConsegna As String
    Dim strDtModifica As String
    Dim strHrModifica As String
    
    With OCT.DS.Fields
        .ResetFields
        .Item("1").NewValue = ID_SOC
        .Item("1").Updated = True
        .Item("2").NewValue = IDEnteEmittente
        .Item("2").Updated = True
        .Item("3").NewValue = Format$(Date, "dd/mm/yyyy")
        .Item("3").Updated = True
        .Item("4").NewValue = Format$(Time, "hh:mm")
        .Item("4").Updated = True
            
        If Len(Trim$(InFile("11").value)) > 0 Then
            strDtConsegna = InFile("11").value
        Else
            strDtConsegna = InFile("10").value
        End If
        strDtConsegna = Right$(strDtConsegna, 2) & _
                        "/" & Mid$(strDtConsegna, 5, 2) & _
                        "/" & Left$(strDtConsegna, 4)
        
        If Len(Trim$(InFile("12").value)) > 0 Then
            strDtModifica = InFile("12").value
            strDtModifica = Right$(strDtModifica, 2) & _
                            "/" & Mid$(strDtModifica, 5, 2) & _
                            "/" & Left$(strDtModifica, 4)
        
            strHrModifica = InFile("13").value
            strHrModifica = Left$(strHrModifica, 2) & ":" & Right$(strHrModifica, 2)
            .Item("5").NewValue = strDtModifica
            .Item("5").Updated = True
            .Item("6").NewValue = strHrModifica
            .Item("6").Updated = True
       End If
        
        .Item("7").NewValue = strDtConsegna
        .Item("7").Updated = True
        .Item("8").NewValue = strDtConsegna
        .Item("8").Updated = True
        .Item("10").NewValue = Trim$(Left$(InFile("3").value, .Item("10").MaxLength))
        .Item("10").Updated = True
        .Item("11").NewValue = Trim$(Left$(InFile("2").value, .Item("11").MaxLength))
        .Item("11").Updated = True
        .Item("12").NewValue = IDCliente
        .Item("12").Updated = True
        .Item("13").NewValue = IDDestinazione
        .Item("13").Updated = True
        Select Case InFile("16").value
            Case 6
                .Item("14").NewValue = "In giornata"
                .Item("14").Updated = True
            Case 7
                .Item("14").NewValue = "Mattino"
                .Item("14").Updated = True
            Case 8
                .Item("14").NewValue = "Pomeriggio"
                .Item("14").Updated = True
        End Select
        IDTipoVendita = "F2"
        .Item("16").NewValue = IDTipoVendita
        .Item("16").Updated = True
        .Item("18").NewValue = Trim$(Left$(InFile("7").value, .Item("18").MaxLength))
        .Item("18").Updated = True
        .Item("21").NewValue = "R"
        .Item("21").Updated = True
    End With
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    ClientKeyUp KeyCode, Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim mobjInField As CTextFileField
    
    Loading = True
    
    ConfigureObjects Me
    mskFileName = gmRegistry.GetSetting("Options", "OrderImportFile", , True)
    mstrUltimoProgressivo = gmRegistry.GetSetting("Options", "PrevOrderImportFileID", "0", True)
    
    Set ErrH = New CErrorHandler
    
    Set AXX = New CParametro
    With AXX.DS
'        Set .DA = DA
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set OCT = New COrdine
    With OCT.DS
'        Set .DA = DA
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set OCR = New COrdineRiga
    With OCR.DS
'        Set .DA = DA
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set ADZ = New CDestinazione
    With ADZ.DS
'        Set .DA = DA
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
'        Set .MultiSoc.DS.DA = DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set AEE = New CEnteEmittente
    With AEE.DS
'        Set .DA = DA
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
'        Set .MultiSoc.DS.DA = DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set ATV = New CTipoVendita
    With ATV.DS
'        Set .DA = DA
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
'        Set .MultiSoc.DS.DA = DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set APR = New CProdotto
    With APR.DS
'        Set .DA = DA
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
'        Set .MultiSoc.DS.DA = DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set APD = New CPosDoganale
    With APD.DS
'        Set .DA = DA
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
'        Set .MultiSoc.DS.DA = DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
    Set ADE = New CDeposito
    With ADE.DS
'        Set .DA = DA
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set AZO = New CZona
    With AZO.DS
'        Set .DA = DA
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
    End With
    
    Set mobjInFile = New CTextFile
'    Select Case gmstrCodInstallazione
'        Case ciInst1
    #If VersionID = 1 Then
            With mobjInFile
                .Add "Tipo Operazione", 1, "20"
                .Add "Zona", 2, "2"
                .Add "Filler1", 2, "3"
                .Add "Codice Destinazione", 8, "4"
                .Add "Ragione Sociale", 30, "5"
                .Add "Indirizzo", 65, "6"
                .Add "Località", 65, "7"
                .Add "CAP", 5, "8"
                .Add "Telefono", 20, "9"
                .Add "Note 1", 50, "10"
                .Add "Note 2", 50, "11"
                .Add "Numero Ordine", 8, "12"
                .Add "Deposito", 2, "13"
                .Add "Descrizione", 30, "14"
                .Add "Km Raggio", 4, "15"
                .Add "DataConsegna", 8, "16"
                .Add "Prodotto", 6, "17"
                .Add "Quantità", 6, "18"
                .Add "Tipo Vendita", 2, "19"
                .Add "Flag Programmato", 1, "21"
                .Add "Flag Urgente", 1, "22"
                .Add "Flag Fascia Consegna", 1, "23"
                ' (PM) 21/07/06 - Aggiunto campo Fecha de Grado e orari
                .Add "Data Grado", 8, "25"
                .Add "Note 3", 50, "24"
                .Add "Orario Apertura 1", 5, "26"
                .Add "Orario Chiusura 1", 5, "27"
                .Add "Orario Apertura 2", 5, "28"
                .Add "Orario Chiusura 2", 5, "29"
                .Add "Orario Apertura 3", 5, "30"
                .Add "Orario Chiusura 3", 5, "31"
            End With
            For i = 0 To 1
                Frame4(i).Visible = (i = 0)
            Next i
            Label1(5).Visible = False
            mskFileNameDettaglio.Visible = False
'        Case ciInst2
        #ElseIf (VersionID = 2) Then 'Or (VersionID = 5) Then
            With mobjInFile
                .Add "Tipo record", 4, "1"
                .Add "Tipo Ordine", 2, "2"
                .Add "Numero Ordine", 9, "3"
                .Add "Ente Emittente", 3, "4"
                .Add "Provincia Destinazione", 2, "5"
                .Add "Progressivo Destinazione", 3, "6"
                .Add "Presidio Logistico", 3, "7"
                .Add "Data Acquisizione", 8, "8"
                .Add "Ora Acquisizione", 4, "9"
                .Add "Data Consegna Auto", 8, "10"
                .Add "Data Consegna Digitata", 8, "11"
                .Add "Data ultima modifica", 8, "12"
                .Add "Ora ultima modifica", 4, "13"
                .Add "Utente ultima modifica", 10, "14"
                .Add "Flag marchio globale", 1, "15"
                .Add "Preferenze Gestore", 1, "16"
                .Add "Quantità minima", 2, "17"
            End With
            Set mobjInFileDettaglio = New CTextFile
            With mobjInFileDettaglio
                .Add "Tipo record", 4, "1"
                .Add "Tipo Ordine", 2, "2"
                .Add "Numero Ordine", 9, "3"
                .Add "Filler", 4, "4"
                .Add "Progressivo riga", 1, "5"
                .Add "Codice Prodotto", 5, "6"
                .Add "Filler", 3, "7"
                .Add "Posizione doganale", 2, "8"
                .Add "Imballo", 3, "9"
                .Add "Quantità richiesta", 5, "10"
                .Add "Quantità modificata", 5, "11"
                .Add "Filler", 37, "12"
            End With
            For i = 0 To 1
                Frame4(i).Visible = (i = 1)
            Next i
            Label1(5).Visible = True
            mskFileNameDettaglio.Visible = True
            mskFileNameDettaglio = gmRegistry.GetSetting("Options", "OrderImportFileDetail", , True)
'        Case ciInst4
        #ElseIf VersionID = 4 Then
            With mobjInFile
                .Add "Data inserimento", 6, "1"
                .Add "Ora inserimento", 6, "2"
                .Add "Codice Destinazione", 6, "3"
                .Add "Qta Prodotto 1", 2, "4"
                .Add "Qta Prodotto 2", 2, "5"
                .Add "Qta Prodotto 3", 2, "6"
                .Add "Qta Prodotto 4", 2, "7"
                .Add "Preferenze Gestore", 1, "8"
            End With
            For i = 0 To 1
                Frame4(i).Visible = (i = 0)
            Next i
            Label1(5).Visible = False
            mskFileNameDettaglio.Visible = False
'    End Select
        #End If
    
    
    TDBGrid1.Columns(2).Visible = False
    'TDBGrid1.Columns(13).Visible = False
    
    Loading = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    gmRegistry.SaveSetting "Options", "OrderImportFile", mskFileName, True
    gmRegistry.SaveSetting "Options", "OrderImportFileDetail", mskFileNameDettaglio, True
    gmRegistry.SaveSetting "Options", "PrevOrderImportFileID", mstrUltimoProgressivo, True
    
    Set XA = Nothing
    Set XALog = Nothing
    Set AXX = Nothing
    Set ADZ = Nothing
    Set ADE = Nothing
    Set AZO = Nothing
    Set OCR = Nothing
    Set OCT = Nothing
    Set APR = Nothing
    Set APD = Nothing
    Set ATV = Nothing
    Set AEE = Nothing
    Set ErrH = Nothing
    Set mobjInFile = Nothing
    Set mobjInFileDettaglio = Nothing
    Set frmOCTImport = Nothing
End Sub







Private Sub mskFileName_Change()
'    Select Case gmstrCodInstallazione
'        Case ciInst1
    #If VersionID = 1 Then
            cmdReadFile.Enabled = (Len(mskFileName) > 0)
'        Case ciInst2
    #ElseIf (VersionID = 2) Then 'Or (VersionID = 5) Then
            cmdReadFile.Enabled = (Len(mskFileName) > 0) And (Len(mskFileNameDettaglio) > 0)
'    End Select
    #End If
    cmdImport.Enabled = False
End Sub


Private Sub mskFileName_GotFocus()
    mintFileInput = 1
End Sub


Private Sub mskFileNameDettaglio_Change()
'    Select Case gmstrCodInstallazione
'        Case ciinst2
    #If (VersionID = 2) Then 'Or (VersionID = 5) Then
            cmdReadFile.Enabled = (Len(mskFileName) > 0) And (Len(mskFileNameDettaglio) > 0)
            cmdImport.Enabled = False
'    End Select
    #End If
End Sub

Private Sub mskFileNameDettaglio_GotFocus()
    mintFileInput = 2
End Sub


