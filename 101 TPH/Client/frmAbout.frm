VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   3630
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   7485
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lsvVersions 
      Height          =   1215
      Left            =   240
      TabIndex        =   7
      Top             =   2280
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   2143
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Modulo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Versione"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Company Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Descrizione"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   480
      Left            =   240
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   480
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5880
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   375
      Left            =   5880
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   1380
      Left            =   4200
      Picture         =   "frmAbout.frx":030A
      Top             =   600
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   90
      X2              =   7320
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "App Description"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   1080
      TabIndex        =   3
      Top             =   240
      Width           =   3795
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Application Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1050
      TabIndex        =   5
      Top             =   480
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   105
      X2              =   5640
      Y1              =   2460
      Y2              =   2460
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      Height          =   225
      Left            =   1080
      TabIndex        =   6
      Top             =   840
      Width           =   3795
   End
   Begin VB.Label lblLicense 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   825
      Left            =   255
      TabIndex        =   4
      Top             =   1200
      Width           =   3630
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub AddVersionListItem(x As ListItem, _
                               FileInfo As CFileVersionInfo, _
                               PathName As String)
    FileInfo.FullPathName = PathName
    Set x = lsvVersions.ListItems.Add(, , FileInfo.OriginalFilename)
    With x
        .SubItems(1) = CStr(FileInfo.FileVersion)
        .SubItems(2) = FileInfo.CompanyName
        .SubItems(3) = FileInfo.FileDescription
    End With
End Sub

Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Dim strLicenseTo As String
    Dim strLicenseName As String
    Dim strPath As String
    Dim x As ListItem
    Dim FileInfo As CFileVersionInfo
    
    RESOIL2MAIN.MousePointer = vbHourglass
    Set FileInfo = New CFileVersionInfo
    
    strPath = CommonFilesDir & "\RES Data\"
    AddVersionListItem x, FileInfo, strPath & "reslib01.dll"
    AddVersionListItem x, FileInfo, strPath & "reslib02.dll"
'    AddVersionListItem x, FileInfo, strPath & "reslib03.dll"
'    AddVersionListItem x, FileInfo, strPath & "reslib05.dll"
    #If VersionID = 2 Then
        AddVersionListItem x, FileInfo, strPath & "oil2blg2.dll"
    #ElseIf VersionID = 4 Then
        AddVersionListItem x, FileInfo, strPath & "oil2blg4.dll"
    #Else
        AddVersionListItem x, FileInfo, strPath & "oil2blg0.dll"
    #End If
    Set FileInfo = Nothing
    GetLicenseInfo strLicenseTo, strLicenseName
    picIcon.Picture = RESOIL2MAIN.Icon
    Me.Caption = "Informazioni su " & App.Title
    lblVersion.Caption = "Versione " & App.major & "." & App.minor & "." & App.Revision
    lblTitle.Caption = App.Title
    lblDescription.Caption = App.CompanyName
    lblLicense.Caption = vbCrLf & Space(3) & strLicenseTo & vbCrLf & Space(3) & strLicenseName
    RESOIL2MAIN.MousePointer = vbDefault
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
    ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "System Information non disponibile", vbOKOnly
End Sub

