VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTrasmAnagrafiche"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public TRAutista As CTextFile
Public TRAutomezzo As CTextFile
Public TRVettore As CTextFile
Public TRDestinazione As CTextFile
Public TRUnitaTrasporto As CTextFile
Public TRAddebiti As CTextFile
Public TRExViagT As CTextFile
Public TRExViagProd As CTextFile

Private Sub Class_Initialize()

    Set TRAutista = New CTextFile
    Set TRAutomezzo = New CTextFile
    Set TRVettore = New CTextFile
    Set TRDestinazione = New CTextFile
    Set TRUnitaTrasporto = New CTextFile
    Set TRAddebiti = New CTextFile
    Set TRExViagT = New CTextFile

    With TRAutista
        .Add "Codice", 6, "1"
        .Add "Nome", 30, "2"
        .Add "Nome Abbreviato", 15, "3"
        .Add "Vettore", 6, "4"
        .Add "Numero Patente", 20, "5"
        .Add "Numero NIF", 16, "6"
        .Add "Disponibilità", 1, "7"
        .Add "Note 1", 40, "8"
        .Add "Note 2", 40, "9"
        .Add "Annullato", 1, "10"
    End With

    With TRDestinazione
        .Add "Codice", 10, "1"
        .Add "Ragione Sociale", 30, "2"
        .Add "Ragione Sociale Aggiuntiva", 30, "3"
        .Add "Indirizzo", 30, "4"
        .Add "Località", 30, "5"
        .Add "Zona statistica", 3, "6"
    End With

    With TRVettore
        .Add "Codice", 6, "1"
        .Add "Ragione Sociale", 40, "2"
        .Add "Ragione Sociale Aggiuntiva", 40, "3"
        .Add "Ragione Sociale Abbreviata", 6, "4"
        .Add "Indirizzo", 40, "5"
        .Add "Località", 25, "6"
        .Add "Provincia", 5, "7"
        .Add "CAP", 10, "8"
        .Add "Codice Fiscale", 16, "9"
'        .Add "Partita IVA", 20, "10"
        .Add "Partita IVA", 12, "10"
'        .Add "Modalità di Pagamento", 3, "11"
        .Add "Modalità di Pagamento", 3, "11"
'        .Add "IVA", 3, "12"
        .Add "IVA", 2, "12"
        .Add "Telefono", 20, "13"
        .Add "Nota 1", 40, "14"
        .Add "Nota 2", 40, "15"
        .Add "Annullato", 1, "16"
    End With

    With TRUnitaTrasporto
        .Add "Codice", 5, "1"
        .Add "Descrizione", 30, "2"
        .Add "Motrice", 10, "3"
        .Add "Rimorchio", 10, "4"
        .Add "Autista 1", 6, "5"
        .Add "Autista 2", 6, "6"
        .Add "Vettore", 6, "7"
        .Add "Vettore Trasportatore", 6, "8"
        .Add "Portata Totale", 5, "9", , fftNumerico
        .Add "Disponibilità", 1, "10"
        .Add "Deposito", 3, "11"
        .Add "Ente Emittente", 3, "12"
        .Add "Società", 6, "13"
        .Add "Sociale/Terzo", 1, "14"
        .Add "Badge", 3, "15"
        .Add "Contratto", 10, "16"
        .Add "Badge", 5, "17"
        .Add "Annullato", 1, "18"
    End With

    With TRAutomezzo
        .Add "Targa", 10, "1"
        .Add "Motrice/Rimorchio", 1, "2"
        .Add "Descrizione", 30, "3"
        .Add "Tipo Automezzo", 5, "4"
        .Add "Disponibilità", 1, "5"
        .Add "Capacità Scomparto 1", 5, "6", , fftNumerico
        .Add "Capacità Scomparto 2", 5, "7", , fftNumerico
        .Add "Capacità Scomparto 3", 5, "8", , fftNumerico
        .Add "Capacità Scomparto 4", 5, "9", , fftNumerico
        .Add "Capacità Scomparto 5", 5, "10", , fftNumerico
        .Add "Capacità Scomparto 6", 5, "11", , fftNumerico
        .Add "Capacità Scomparto 7", 5, "12", , fftNumerico
        .Add "Capacità Scomparto 8", 5, "13", , fftNumerico
        .Add "Capacità Scomparto 9", 5, "14", , fftNumerico
        .Add "Capacità Scomparto 10", 5, "15", , fftNumerico
        .Add "Numero Scomparti 1", 2, "16", , fftNumerico
        .Add "Numero Scomparti 2", 2, "17", , fftNumerico
        .Add "Numero Scomparti 3", 2, "18", , fftNumerico
        .Add "Numero Scomparti 4", 2, "19", , fftNumerico
        .Add "Numero Scomparti 5", 2, "20", , fftNumerico
        .Add "Numero Scomparti 6", 2, "21", , fftNumerico
        .Add "Numero Scomparti 7", 2, "22", , fftNumerico
        .Add "Numero Scomparti 8", 2, "23", , fftNumerico
        .Add "Numero Scomparti 9", 2, "24", , fftNumerico
        .Add "Numero Scomparti 10", 2, "25", , fftNumerico
        .Add "Portata", 6, "26", , fftNumerico
        .Add "Tipo Cisterna", 3, "27"
        .Add "Tara Pesata", 6, "28", , fftNumerico
        .Add "Tara Libretto", 6, "29", , fftNumerico
        .Add "Portata rimorchiabile", 6, "30", , fftNumerico
        .Add "Cubatura costruttore", 6, "31", , fftNumerico
        .Add "Tipo Carica", 1, "32"
        .Add "Satellitare installato", 2, "33"
        .Add "Annullato", 1, "34"
    End With

    With TRAddebiti
        .Add "Anno", 4, "1"
        .Add "Filiale", 2, "2"
        .Add "Chiave", 7, "3"
        .Add "Ctr", 3, "4"
        .Add "Voce Addebito Cliente", 3, "5"
        .Add "Descrizione Addebito", 30, "6"
        .Add "Filiale Cliente", 2, "7"
        .Add "Codice Cliente", 6, "8"
        .Add "Contr. Fatturazione", 3, "9"
        .Add "Unità Misura", 2, "10"
        .Add "Quantità", 11, "11"
        .Add "Tariffa", 13, "12"
        .Add "Importo", 15, "13"
        .Add "Codice IVA", 2, "14"
        .Add "Anno Fattura", 4, "15"
        .Add "Filiale Emissione", 2, "16"
        .Add "Numero Fattura", 7, "17"
        .Add "Data Fattura", 8, "18"
        .Add "Rif. Anno Documento", 4, "19"
        .Add "Rif. Filiale Documento", 2, "20"
        .Add "Rif. Numero Documento", 7, "21"
        .Add "Causale Documento", 3, "22"
        .Add "Pagamento", 3, "23"
        .Add "Codice Divisa", 3, "24"
        .Add "Cambio Euro", 11, "25"
        .Add "Data Viaggio", 8, "26"
        .Add "Filiale Fatturazione", 2, "27"
        .Add "Tipo Documento", 1, "28"
        .Add "Conto", 30, "29"
        .Add "Resoil Key 1", 15, "30"
        .Add "Resoil Key 2", 4, "31"
        .Add "Annullamento", 1, "32"
    End With
    
    With TRExViagT
        .Add "Anno viaggio", 4, "1", , fftNumerico
        .Add "Filiale viaggio", 2, "2", "01", fftAlfa
        .Add "Key viaggio", 7, "3", , fftNumerico
        .Add "Ctr viaggio", 3, "4", , fftNumerico
        .Add "Fil. cliente", 2, "5", "01", fftAlfa
        .Add "Codice cliente", 6, "6", , fftNumerico
        .Add "Contratto cliente", 3, "7", , fftNumerico
        .Add "Codice partenza", 5, "8", , fftAlfa
        .Add "Località partenza", 30, "9", , fftAlfa
        .Add "Codice destino", 5, "10", , fftAlfa
        .Add "Località destino", 30, "11", , fftAlfa
        .Add "Fil. vettore", 2, "12", "01", fftAlfa
        .Add "Codice vettore", 6, "13", , fftNumerico
        .Add "Contratto vettore", 3, "14", , fftNumerico
        .Add "Data servizio", 8, "15", , fftNumerico
        .Add "Trattore", 10, "16", , fftAlfa
        .Add "Rimorchio", 10, "17", , fftAlfa
        .Add "Tipo merce", 3, "18", , fftAlfa
        .Add "Maggiorazione sabato cliente", 1, "19", "S", fftAlfa
        .Add "Maggiorazione festivo cliente", 1, "20", "F", fftAlfa
        .Add "Km tachigrafo", 6, "21", "0", fftNumerico
        .Add "Km viaggio", 6, "22", , fftNumerico
        .Add "Peso", 7, "23", , fftNumerico
        .Add "N° ddt", 15, "24", , fftAlfa
        .Add "Data ddt", 8, "25", , fftNumerico
        .Add "Unità viaggiante", 4, "26", , fftAlfa
        .Add "Maggiorazione bomba vettore", 1, "27", " ", fftAlfa
        .Add "Maggiorazione festivo vettore", 1, "28", "F", fftAlfa
        .Add "Sconto ads", 1, "29", "1", fftAlfa
        .Add "Consegne", 2, "30", , fftNumerico
        .Add "Filiale deposito", 2, "31", "01"
        .Add "Codice deposito", 6, "32", , fftNumerico
        .Add "Progressivo", 3, "33", "0", fftNumerico
        .Add "Filiale destinatario", 2, "34", "01"
        .Add "Codice destinatario", 6, "35", , fftNumerico
        .Add "Progressivo", 3, "36", "0", fftNumerico
        .Add "Numero ordine", 15, "37"
        .Add "Numero ordine di carico", 15, "38"
        .Add "Area AGIP", 3, "39"
        .Add "Terzo riferimento", 15, "40"
        .Add "Tipo servizio", 3, "41"
        .Add "s/compte", 30, "42"
        .Add "Mittente", 10, "43"
        .Add "Destinazione", 10, "44"
        ' (PM) - Raggruppamento
        .Add "Raggruppamento", 3, "45"
        ' (PM) - Km Passivi
        .Add "Km passivi", 6, "46", , fftNumerico
    End With

    Set TRExViagProd = New CTextFile
    With TRExViagProd
        .Add "Anno viaggio", 4, "1", , fftNumerico
        .Add "Filiale viaggio", 2, "2", "01", fftAlfa
        .Add "Key viaggio", 7, "3", , fftNumerico
        .Add "Ctr viaggio", 3, "4", , fftNumerico
        .Add "Progressivo", 3, "5", "0", fftNumerico
        .Add "Tipo prodotto", 3, "6"
        .Add "Quantità", 6, "7", , fftNumerico
    End With

End Sub

Private Sub Class_Terminate()
    Set TRAutista = Nothing
    Set TRAutomezzo = Nothing
    Set TRVettore = Nothing
    Set TRDestinazione = Nothing
    Set TRUnitaTrasporto = Nothing
    Set TRAddebiti = Nothing
    Set TRExViagT = Nothing
    Set TRExViagProd = Nothing
End Sub


