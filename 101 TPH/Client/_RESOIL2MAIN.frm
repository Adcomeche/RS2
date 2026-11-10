VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm RESOIL2MAIN 
   BackColor       =   &H00FFFFFF&
   Caption         =   "RES/OIL.2"
   ClientHeight    =   3570
   ClientLeft      =   165
   ClientTop       =   1110
   ClientWidth     =   7935
   Icon            =   "RESOIL2MAIN.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   3195
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Text            =   "RES DATA s.r.l."
            TextSave        =   "RES DATA s.r.l."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10927
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuL1 
      Caption         =   "&1 Anagrafiche"
      Index           =   1
      Begin VB.Menu mnuL21 
         Caption         =   "1.&1 Società"
         Index           =   1
      End
      Begin VB.Menu mnuL21 
         Caption         =   "1.&2 Enti Emittenti"
         Index           =   2
      End
      Begin VB.Menu mnuL21 
         Caption         =   "1.&3 Depositi"
         Index           =   3
      End
      Begin VB.Menu mnuL21 
         Caption         =   "1.&4 Clienti"
         Index           =   4
      End
      Begin VB.Menu mnuL21 
         Caption         =   "1.&5 Destinazioni"
         Index           =   5
      End
      Begin VB.Menu mnuL21 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuL21 
         Caption         =   "1.&6 Gestione Prodotti"
         Index           =   7
         Begin VB.Menu mnuL317 
            Caption         =   "1.6.&1 Prodotti"
            Index           =   1
         End
         Begin VB.Menu mnuL317 
            Caption         =   "1.6.&2 Imballi"
            Index           =   2
         End
         Begin VB.Menu mnuL317 
            Caption         =   "1.6.&3 Posizioni Doganali"
            Index           =   3
         End
      End
      Begin VB.Menu mnuL21 
         Caption         =   "1.&7 Gestione Trasporti"
         Index           =   8
         Begin VB.Menu mnuL318 
            Caption         =   "1.7.&1 Vettori"
            Index           =   1
         End
         Begin VB.Menu mnuL318 
            Caption         =   "1.7.&2 Autisti"
            Index           =   2
         End
         Begin VB.Menu mnuL318 
            Caption         =   "1.7.&3 Automezzi"
            Index           =   3
         End
         Begin VB.Menu mnuL318 
            Caption         =   "1.7.&4 Unità di Trasporto"
            Index           =   4
         End
         Begin VB.Menu mnuL318 
            Caption         =   "1.7.&5 Contratti"
            Index           =   5
         End
      End
      Begin VB.Menu mnuL21 
         Caption         =   "1.&8 Gestione Tariffe"
         Index           =   9
         Begin VB.Menu mnuL319 
            Caption         =   "1.8.&1 Fasce Tariffarie"
            Index           =   1
         End
      End
      Begin VB.Menu mnuL21 
         Caption         =   "1.&9 Gestione Tabelle"
         Index           =   10
         Begin VB.Menu mnuL3110 
            Caption         =   "1.9.&1 Modalità di Pagamento"
            Index           =   1
         End
         Begin VB.Menu mnuL3110 
            Caption         =   "1.9.&2 IVA"
            Index           =   2
         End
         Begin VB.Menu mnuL3110 
            Caption         =   "1.9.&3 Zone"
            Index           =   3
         End
         Begin VB.Menu mnuL3110 
            Caption         =   "1.9.&4 Valute"
            Index           =   4
         End
         Begin VB.Menu mnuL3110 
            Caption         =   "1.9.&5 Tipologie"
            Index           =   5
            Begin VB.Menu mnuL4195 
               Caption         =   "1.9.5.&1 Tipi Cliente"
               Index           =   1
            End
            Begin VB.Menu mnuL4195 
               Caption         =   "1.9.5.&2 Tipi Destinazione"
               Index           =   2
            End
            Begin VB.Menu mnuL4195 
               Caption         =   "1.9.5.&3 Tipi Automezzo"
               Index           =   3
            End
            Begin VB.Menu mnuL4195 
               Caption         =   "1.9.5.&4 Tipi Cisterna"
               Index           =   4
            End
            Begin VB.Menu mnuL4195 
               Caption         =   "1.9.5.&5 Tipi Vendita"
               Index           =   5
            End
            Begin VB.Menu mnuL4195 
               Caption         =   "1.9.5.&6 Tipi Contratto"
               Index           =   6
            End
         End
         Begin VB.Menu mnuL3110 
            Caption         =   "1.9.&6 Raggruppamenti"
            Index           =   6
            Begin VB.Menu mnuL4196 
               Caption         =   "1.9.6.&1 Raggruppamenti Cliente"
               Index           =   1
            End
            Begin VB.Menu mnuL4196 
               Caption         =   "1.9.6.&2 Raggruppamenti Destinazione"
               Index           =   2
            End
            Begin VB.Menu mnuL4196 
               Caption         =   "1.9.6.&3 Raggruppamenti Deposito"
               Index           =   3
            End
            Begin VB.Menu mnuL4196 
               Caption         =   "1.9.6.&4 Raggruppamenti Prodotto"
               Index           =   4
            End
            Begin VB.Menu mnuL4196 
               Caption         =   "1.9.6.&5 Raggruppamenti Ordine"
               Index           =   5
            End
            Begin VB.Menu mnuL4196 
               Caption         =   "1.9.6.&6 Raggruppamenti Fiscali Prodotto"
               Index           =   6
            End
         End
         Begin VB.Menu mnuL3110 
            Caption         =   "1.9.&7 Tabelle di Sistema"
            Index           =   7
            Begin VB.Menu mnuL4197 
               Caption         =   "1.9.7.&1 Parametri di Sistema"
               Index           =   1
            End
            Begin VB.Menu mnuL4197 
               Caption         =   "1.9.7.&2 Numeratori"
               Index           =   2
            End
            Begin VB.Menu mnuL4197 
               Caption         =   "1.9.7.&3 Parametri Tariffe"
               Index           =   3
            End
         End
      End
   End
   Begin VB.Menu mnuL1 
      Caption         =   "&2 Ordini"
      Index           =   2
      Begin VB.Menu mnuL22 
         Caption         =   "2.&1 Inserimento"
         Index           =   1
      End
      Begin VB.Menu mnuL22 
         Caption         =   "2.&2 Visualizzazioni"
         Index           =   2
         Begin VB.Menu mnuL322 
            Caption         =   "2.2.&1 Visualizzazione Ordini da eseguire"
            Index           =   1
         End
         Begin VB.Menu mnuL322 
            Caption         =   "2.2.&2 Visualizzazione Ordini in esecuzione"
            Index           =   2
         End
         Begin VB.Menu mnuL322 
            Caption         =   "2.2.&3 Visualizzazione Ordini pendenti"
            Index           =   3
         End
      End
      Begin VB.Menu mnuL22 
         Caption         =   "2.&3 Situazione Ordini"
         Index           =   3
      End
      Begin VB.Menu mnuL22 
         Caption         =   "2.&4 Evasione Ordini"
         Index           =   4
      End
   End
   Begin VB.Menu mnuL1 
      Caption         =   "&3 Viaggi"
      Index           =   3
      Begin VB.Menu mnuL23 
         Caption         =   "3.&1 Inserimento"
         Index           =   1
      End
      Begin VB.Menu mnuL23 
         Caption         =   "3.&2 Modifica Viaggi provvisori"
         Index           =   2
      End
      Begin VB.Menu mnuL23 
         Caption         =   "3.&3 Visualizzazioni"
         Index           =   3
         Begin VB.Menu mnuL313 
            Caption         =   "3.3.&1 Visualizzazione Viaggi definitivi"
            Index           =   1
         End
         Begin VB.Menu mnuL313 
            Caption         =   "3.3.&2 Visualizzazione Viaggi trasmessi"
            Index           =   2
         End
         Begin VB.Menu mnuL313 
            Caption         =   "3.3.&3 Visualizzazione Viaggi completati"
            Index           =   3
         End
      End
      Begin VB.Menu mnuL23 
         Caption         =   "3.&4 Completamento Viaggi"
         Index           =   4
      End
      Begin VB.Menu mnuL23 
         Caption         =   "3.&5 Situazione Viaggi"
         Index           =   5
      End
   End
   Begin VB.Menu mnuL1 
      Caption         =   "&4 Prefatturazione"
      Index           =   4
      Begin VB.Menu mnuL24 
         Caption         =   "4.&1 Archivio storico proposte di fattura"
         Index           =   1
      End
      Begin VB.Menu mnuL24 
         Caption         =   "4.&2 Valorizzazione delle Consegne"
         Index           =   2
      End
      Begin VB.Menu mnuL24 
         Caption         =   "4.&3 Valorizzazione degli Ordini"
         Index           =   3
      End
      Begin VB.Menu mnuL24 
         Caption         =   "4.&4 Gestione prefatture"
         Index           =   4
      End
   End
   Begin VB.Menu mnuL1 
      Caption         =   "&5 Funzioni Statistiche"
      Index           =   5
      Begin VB.Menu mnuL25 
         Caption         =   "5.&1 Interrogazione"
         Index           =   1
      End
   End
   Begin VB.Menu mnuL1 
      Caption         =   "&6 Funzioni di Interfaccia"
      Index           =   6
      Begin VB.Menu mnuL26 
         Caption         =   "6.&1 Ricezione Ordini"
         Index           =   1
      End
      Begin VB.Menu mnuL26 
         Caption         =   "6.&2 Trasferimento Ordini a Mover"
         Index           =   2
      End
      Begin VB.Menu mnuL26 
         Caption         =   "6.&3 Ricezione Viaggi da Mover"
         Index           =   3
      End
      Begin VB.Menu mnuL26 
         Caption         =   "6.&4 Trasmissione Viaggi"
         Index           =   4
      End
      Begin VB.Menu mnuL26 
         Caption         =   "6.&5 Trasmissione Viaggi a Dep. Comune"
         Index           =   5
      End
      Begin VB.Menu mnuL26 
         Caption         =   "6.&6 Ricezione Anagrafiche"
         Index           =   6
      End
      Begin VB.Menu mnuL26 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuL26 
         Caption         =   "6.&7 Ricezione Proposte di Fattura"
         Index           =   8
      End
      Begin VB.Menu mnuL26 
         Caption         =   "6.&8 Trasferimento Destinazioni a Mapshort"
         Index           =   9
      End
      Begin VB.Menu mnuL26 
         Caption         =   "6.&9 Ricezione distanze da Mapshort"
         Index           =   10
      End
   End
   Begin VB.Menu mnuL1 
      Caption         =   "&7 Strumenti"
      Index           =   7
      Begin VB.Menu mnuL27 
         Caption         =   "7.&1 Opzioni"
         Index           =   1
      End
      Begin VB.Menu mnuL27 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuL27 
         Caption         =   "7.&2 Informazioni su RES/OIL.2"
         Index           =   3
      End
      Begin VB.Menu mnuL27 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuL27 
         Caption         =   "7.&9 Esci"
         Index           =   5
      End
   End
   Begin VB.Menu mnuL1 
      Caption         =   "&8 Finestra"
      Index           =   8
      WindowList      =   -1  'True
      Begin VB.Menu mnuL28 
         Caption         =   "8.1 Sovrapponi"
         Index           =   1
      End
      Begin VB.Menu mnuL28 
         Caption         =   "8.2 Disponi icone"
         Index           =   2
      End
      Begin VB.Menu mnuL28 
         Caption         =   "8.3 Riduci a icona"
         Index           =   3
      End
   End
End
Attribute VB_Name = "RESOIL2MAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Unload(Cancel As Integer)
    CleanUp
End Sub

Private Sub mnuL21_Click(Index As Integer)
    ShowWindow mnuL21(Index)
End Sub


Private Sub mnuL22_Click(Index As Integer)
    ShowWindow mnuL22(Index)
End Sub

Private Sub mnuL23_Click(Index As Integer)
    ShowWindow mnuL23(Index)
End Sub

Private Sub mnuL24_Click(Index As Integer)
    ShowWindow mnuL24(Index)
End Sub


Private Sub mnuL25_Click(Index As Integer)
    ShowWindow mnuL25(Index)
End Sub

Private Sub mnuL26_Click(Index As Integer)
    ShowWindow mnuL26(Index)
End Sub

Private Sub mnuL27_Click(Index As Integer)
    Select Case Index
        Case 1
            frmOptions.Show
        Case 3
            frmAbout.Show
        Case 5
            Unload Me
    End Select
End Sub

Private Sub mnuL28_Click(Index As Integer)
    Dim F As Form
    
    Select Case Index
        Case 1
            On Error GoTo NextF
            
            For Each F In Forms
                If F.MDIChild Then
                    F.WindowState = vbNormal
                End If
NextF:
            Next F
            Me.Arrange vbCascade
        Case 2
            Me.Arrange vbArrangeIcons
        Case 3
            On Error GoTo NextF1
            
            For Each F In Forms
                If F.MDIChild Then
                    F.WindowState = vbMinimized
                End If
NextF1:
            Next F
    End Select
End Sub

Private Sub mnuL313_Click(Index As Integer)
    ShowWindow mnuL313(Index)

End Sub

Private Sub mnuL317_Click(Index As Integer)
    ShowWindow mnuL317(Index)
End Sub
Private Sub mnuL3110_Click(Index As Integer)
    ShowWindow mnuL3110(Index)
End Sub

Private Sub mnuL318_Click(Index As Integer)
    ShowWindow mnuL318(Index)
End Sub

Private Sub mnuL319_Click(Index As Integer)
    ShowWindow mnuL319(Index)
End Sub

Private Sub mnuL322_Click(Index As Integer)
    ShowWindow mnuL322(Index)
End Sub




Private Sub mnuL4195_Click(Index As Integer)
    ShowWindow mnuL4195(Index)
End Sub

Private Sub mnuL4196_Click(Index As Integer)
    ShowWindow mnuL4196(Index)

End Sub


Private Sub mnuL4197_Click(Index As Integer)
    ShowWindow mnuL4197(Index)
End Sub


