VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIncidencias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Completar Incidencias"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   8145
   Begin VB.TextBox txt_cantidad 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CheckBox chk_medprop 
      Caption         =   "Resuelto Medios Propios"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5040
      TabIndex        =   5
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox txt_timeparaliz 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   2760
      Width           =   2175
   End
   Begin VB.CommandButton btn_envia 
      Caption         =   "Envia"
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   3240
      Width           =   1695
   End
   Begin VB.ComboBox cmb_motivoEvento 
      Height          =   315
      Left            =   2760
      TabIndex        =   2
      Top             =   240
      Width           =   4815
   End
   Begin VB.ComboBox cmb_prod1 
      Height          =   315
      Left            =   2760
      TabIndex        =   1
      Top             =   1800
      Width           =   1815
   End
   Begin VB.ComboBox cmb_prod2 
      Height          =   315
      Left            =   2760
      TabIndex        =   0
      Top             =   2280
      Width           =   1815
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   3945
      Width           =   8145
      _ExtentX        =   14367
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
            Object.Width           =   8916
            Key             =   "Msg"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "10:19"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "12/12/2011"
            Key             =   "Date"
         EndProperty
      EndProperty
   End
   Begin MSMask.MaskEdBox mskField6 
      Height          =   255
      Left            =   840
      TabIndex        =   13
      Tag             =   "13"
      Top             =   720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   12648447
      MaxLength       =   25
      PromptChar      =   "_"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7560
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
            Picture         =   "frmIncidencias.frx":0000
            Key             =   "GreenCheck"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIncidencias.frx":0112
            Key             =   "RedCross"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIncidencias.frx":0224
            Key             =   "Editing"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblField2 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2760
      TabIndex        =   15
      Tag             =   "13"
      Top             =   720
      Width           =   4935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "&Cliente"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   14
      Top             =   720
      Width           =   735
   End
   Begin VB.Image imgState 
      Height          =   240
      Index           =   3
      Left            =   7800
      Tag             =   "13"
      Top             =   720
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Cantidad Involucrada en el evento: "
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Cod. Producto 1"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Cod. Producto 2"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Tiempo Paralización (en Horas)"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "Motivo Evento: "
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmIncidencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public WithEvents DataClass As COrdine
Attribute DataClass.VB_VarHelpID = -1
Private ADZ As CDestinazione

Public State As FormState
Public Loading As Boolean
Public mblnConcurError As Boolean
Public Soc As CSocieta


Public Sub SetState(NewState As FormState)
    DataClass.DS.SetClientState NewState
'    Select Case NewState
'        Case fsNone
'            If Not Loading Then
'                cmb_motivoEvento.SetFocus
'            End If
'        Case fsEdit
'            cmb_motivoEvento.SetFocus
'        Case fsInsert
'           cmb_motivoEvento.SetFocus
'    End Select
End Sub



Private Sub acitvarForm(idEvento As String)
 
Label1(3).Visible = False
mskField6.Visible = False
lblField2.Visible = False

Select Case idEvento
        Case "009"
        Case "010"
            txt_cantidad.Enabled = True
        Case "011"
            txt_cantidad.Enabled = True
        Case "012"
            txt_cantidad.Enabled = True
            chk_medprop.Enabled = True
        Case "013"
        Case "014"
        Case "015"
        Case "016"
            txt_cantidad.Enabled = True
        Case "017"
            Label1(3).Visible = True
            mskField6.Visible = True
            lblField2.Visible = True
            txt_cantidad.Enabled = True
        Case "018"
            txt_cantidad.Enabled = True
        Case "019"
            txt_timeparaliz.Enabled = True
        Case "020"
            txt_timeparaliz.Enabled = True
End Select

End Sub

Private Function checkForm(idEvento As String) As Boolean
     Dim bueno As Boolean
     bueno = True
    
    Select Case idEvento
        Case "009"
        Case "010"
           If txt_cantidad.Text = "" Then
                bueno = False
           End If
        Case "011"
            If txt_cantidad.Text = "" Then
                bueno = False
            End If
        Case "012"
            If txt_cantidad.Text = "" Then
                bueno = False
            End If
        Case "013"
        Case "014"
        Case "015"
        Case "016"
            If txt_cantidad.Text = "" Then
                bueno = False
            End If
        Case "017"
            If mskField6.Text = "" Then
                bueno = False
            End If
            If txt_cantidad.Text = "" Then
                bueno = False
            End If
        Case "018"
            If txt_cantidad.Text = "" Then
                bueno = False
            End If
        Case "019"
            If txt_timeparaliz.Text = "" Or Not IsNumeric(txt_timeparaliz.Text) Then
                bueno = False
            End If
        Case "020"
            If txt_timeparaliz.Text = "" Or Not IsNumeric(txt_timeparaliz.Text) Then
                bueno = False
            End If
        
End Select

    
checkForm = bueno
End Function

Private Sub btn_envia_Click()

    Dim DA As CDataAccess
    Set DA = DataConnect.DA
    Dim HREV, DYEV As String
    Dim strIDViaggio As String
    Dim strIDConsegna As String
    Dim CodEv As String
    Dim COD, CODINC, codentrega, SQL_GEN_EV As String
    Dim BKM As Variant
    Dim Respuesta As VbMsgBoxResult
    BKM = frmVVTSnapshotExport.TDBGrid1.bookmark
    strIDViaggio = frmVVTSnapshotExport.txt_idviag
    COD = frmVVTSnapshotExport.txt_cod
    Dim exito, FechaOk As Boolean
    exito = False
    FechaOk = True
    HREV = frmVVTSnapshotExport.txt_hrev.Text
    DYEV = frmVVTSnapshotExport.txt_dataev.Text
    CodEv = frmVVTSnapshotExport.txt_CodEv
    
    'VarInc
    Dim CantidadINC As String
    Dim Motivo As String
    Dim codpro1 As String
    Dim codpro2 As String
    Dim medprop As String
    Dim Tiempo As String
    Dim DestinoCep As String
    
    CantidadINC = txt_cantidad.Text
    Motivo = Left(cmb_motivoEvento, 3)
     Tiempo = txt_timeparaliz.Text
    codpro1 = cmb_prod1
    codpro2 = cmb_prod2
    medprop = chk_medprop.value

    If CodEv = "017" Then
        
         With ADZ.DS
                     .Fields.ResetFields
                      .Fields("0").FindValue = mskField6.Text
                    If (Not ADZ.Find) Or (.HowMany = 0) Then
                    MsgBox ("El codigo cliente no es correcto")
                     End If
         '            DestinoCep = .MultiSoc.DS.GetFirmCode(ID_SOC_CEPSA)
         End With
    
        DestinoCep = Right(mskField6.Text, 8)
    End If
   
   If checkForm(CodEv) Then

    
            Respuesta = MsgBox("Confirmación." + Chr(13) + " Este evento genera una incidencia en Cepsa, " + Chr(13) + "¿Está seguro de enviar la incidencia?", vbYesNo, "Confirmación")
                         
            If Respuesta = vbYes Then
                             
                 SQL_GEN_EV = "INSERT INTO EventosMBFlotasManual (CodigoServicio,CodigoTarea,CodigoEvento,FechaEvento,FechaServidor,EntaEnLoc,MotivoINC, CantidadInvolucrada, CodCLHprod1, CodCLHprod2, ResueltoMediosPropios, TiempoParado, DestinoCep) VALUES ('" + strIDViaggio + "','" + COD + "','" + CodEv + "','" + DYEV + " " + HREV + "',GETDATE (),'0','" + Motivo + "', '" + CantidadINC + "', '" + codpro1 + "', '" + codpro2 + "'," + medprop + ", '" + Tiempo + "', '" + DestinoCep + "')"
                 If DA.ExecuteQuery(SQL_GEN_EV) Then
                    MsgBox ("Evento '" + CodEv + "' generado con exito   El: " + CStr(DYEV) + " A las :" + CStr(HREV))
                    frmVVTSnapshotExport.txt_hrev.Text = "__:__"
                    frmVVTSnapshotExport.txt_dataev.Text = "__/__/__"
                    exito = True
                 Else
                    MsgBox ("")
                 End If
                       
            Else
               MsgBox ("Incidencia NO generada")
            End If
                         
                         
    Unload Me
     
   Else
        
           MsgBox ("Faltan campos por rellenar")
   End If
   

End Sub


Private Sub Form_Load()

Dim Varkbuffer As Variant
Dim Varkbuffertip As Variant
Dim lngHowMany As Long
Dim lngHowManytip As Long
Dim i As Integer
Dim SQL_Productos As String
Dim SQL_Tipologia As String

acitvarForm (frmVVTSnapshotExport.txt_CodEv)
Loading = True
Dim DA As CDataAccess
Set DA = DataConnect.DA
Dim DS As New CDataStruct
Set ErrH = New CErrorHandler
Set DataClass = New COrdine
    
    With DataClass.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        Set .FormHandle = Me
        .FormType = ftEdtRcd1
        .ConfigureInterface
        ConfigureObjects Me
    End With
    
    Set ADZ = New CDestinazione
    With ADZ.DS
        Set .DA = DataConnect.DA
        Set .ErrH = ErrH
        Set .MultiSoc.DS.DA = .DA
        Set .MultiSoc.DS.ErrH = ErrH
    End With
    
'CARGAMOS COMBO PRODUCTOS
   SQL_Productos = "Select APRIDAPR FROM ANAGAPR0 ORDER BY APRIDAPR"
   
   If Not DA.QueryCommand(SQL_Productos, Varkbuffer, lngHowMany) Then
        objLogger.logError "Query SQL_Productos"
   End If

    cmb_prod1.AddItem ("")
    cmb_prod2.AddItem ("")

 For i = 0 To lngHowMany - 1
    cmb_prod1.AddItem (Varkbuffer(0, i))
    cmb_prod2.AddItem (Varkbuffer(0, i))
 
 Next i
    
    
'CARGAMOS COMBO MOTIVOS
                             
    SQL_Tipologia = "Select INCIDCOD, INCDESC0  FROM ANAGINC0"
   
   If Not DA.QueryCommand(SQL_Tipologia, Varkbuffertip, lngHowManytip) Then
        objLogger.logError "Query SQL_Productos"
   End If

    cmb_motivoEvento.AddItem ("")

 For i = 0 To lngHowManytip - 1
    cmb_motivoEvento.AddItem (Varkbuffertip(0, i) + "-" + Varkbuffertip(1, i))
 Next i
                                                      
                                                   
Loading = False
End Sub
