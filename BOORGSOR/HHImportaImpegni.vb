Imports System.Data
Imports System.IO
Imports ExcelDataReader
Imports NTSInformatica.CLN__STD
Imports System.Globalization

'Dati minimi estratti da una riga del file Excel.
Friend Class HH_RigaImpegnoImport
    Public Property TipoRiga As String
    Public Property Codart As String
    Public Property Descrizione As String
    Public Property Quantita As Decimal
End Class

'Raggruppamento preventivo delle righe appartenenti allo stesso impegno.
Friend Class HH_DocumentoImpegnoImport
    Public Property Riferimento As String
    Public Property Cliente As String
    Public Property CodDest As Integer
    Public Property DataBaseConsegna As Date
    Public ReadOnly Property Righe As New List(Of HH_RigaImpegnoImport)
End Class

'Esito sintetico mostrato al termine dell'importazione.
Friend Class HH_EsitoImportazioneImpegni
  Public Property DocumentiSalvati As Integer
  Public Property DocumentiSaltati As Integer
    Public ReadOnly Property ErroriDocumento As New List(Of String)
End Class

Friend Class HH_ImportatoreImpegni

    Public Event Avanzamento(ByVal messaggio As String, ByVal percentuale As Integer)

    Private ReadOnly _oApp As CLE__APP
    Private ReadOnly _oMenu As CLE__MENU
    Private ReadOnly _oClfGsor As CLFORGSOR
    Private ReadOnly _clienti As Dictionary(Of String, Integer)
    Private _oCleGsor As CLEORGSOR
    Private _ultimoMessaggioEntity As String = ""
    Private _ultimoErrore As String = ""

    Public Sub New(ByVal oApp As CLE__APP, ByVal oMenu As CLE__MENU,
                 ByVal oClfGsor As CLFORGSOR)
        Try
            _oApp = oApp
            _oMenu = oMenu
            _oClfGsor = oClfGsor

            'Aggiungere qui le nuove associazioni cliente/conto.
            _clienti = New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase) From {
        {"Iris Mobili S.r.l.", 2020018}
      }
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
        End Try
    End Sub

#Region "Flusso importazione"

    Public Function Importa(ByVal percorsoFile As String) As HH_EsitoImportazioneImpegni
        Dim esito As New HH_EsitoImportazioneImpegni()
        Try
            MostraAvanzamento("Lettura del file Excel...", 0)
            Dim documenti As List(Of HH_DocumentoImpegnoImport) = LeggiERaggruppa(percorsoFile)
            MostraAvanzamento("File letto: " & documenti.Count.ToString() & " impegni da elaborare.", 5)

            'La stessa BEORGSOR viene inizializzata una volta e riutilizzata in sequenza.
            If Not InizializzaEntityOrdini() Then
                esito.ErroriDocumento.Add("Inizializzazione: " & _ultimoErrore)
                Return esito
            End If

            For indiceDocumento As Integer = 0 To documenti.Count - 1
                Dim documento As HH_DocumentoImpegnoImport = documenti(indiceDocumento)
                Dim esistenti As DataTable = _oClfGsor.GetOrdiniPerRiferimento(documento.Riferimento)
                If esistenti Is Nothing Then
                    esito.ErroriDocumento.Add(documento.Riferimento &
                                              ": verifica degli ordini esistenti non riuscita")
                    Continue For
                End If
                If esistenti.Rows.Count > 0 AndAlso
                   Not ConfermaReimportazione(documento.Riferimento, esistenti) Then
                    esito.DocumentiSaltati += 1
                    Continue For
                End If
                Dim motivo As String = ""
                If CreaESalvaDocumento(documento, indiceDocumento + 1, documenti.Count, motivo) Then
                    esito.DocumentiSalvati += 1
                Else
                    esito.ErroriDocumento.Add(documento.Riferimento & ": " & motivo)
                End If
            Next

            MostraAvanzamento("Importazione impegni completata.", 100)

        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
            esito.ErroriDocumento.Add("File: " & ex.Message)
        End Try
        Return esito
    End Function

    Private Function CreaESalvaDocumento(ByVal documento As HH_DocumentoImpegnoImport,
                        ByVal indiceDocumento As Integer,
                        ByVal totaleDocumenti As Integer,
                        ByRef motivo As String) As Boolean
        Try
            MostraAvanzamento("Impegno " & indiceDocumento.ToString() & " di " &
        totaleDocumenti.ToString() & " (" & documento.Riferimento & "): creazione...",
        PercentualeDocumento(indiceDocumento, totaleDocumenti, 0, documento.Righe.Count))

            Dim conto As Integer = 0
            If Not _clienti.TryGetValue(documento.Cliente.Trim(), conto) Then
                motivo = "cliente non configurato: " & documento.Cliente
                Return False
            End If

            Dim noteNonBloccanti As New List(Of String)
            Dim codDest As Integer = documento.CodDest
            Dim dataConsegna As Nullable(Of Date) = Nothing
            RisolviDestinazione(conto, documento.CodDest, documento.DataBaseConsegna,
                                codDest, dataConsegna, noteNonBloccanti)

            If Not NuovoImpegno(conto, documento.Riferimento, codDest, dataConsegna) Then
                motivo = _ultimoErrore
                Return False
            End If

            AggiungiAvvisoEntity(noteNonBloccanti)
            For indiceRiga As Integer = 0 To documento.Righe.Count - 1
                Dim riga As HH_RigaImpegnoImport = documento.Righe(indiceRiga)
                If Not InserisciRiga(riga, noteNonBloccanti) Then
                    motivo = _ultimoErrore
                    Return False
                End If
                MostraAvanzamento("Impegno " & indiceDocumento.ToString() & " di " &
            totaleDocumenti.ToString() & " (" & documento.Riferimento & "): riga " &
            (indiceRiga + 1).ToString() & " di " & documento.Righe.Count.ToString(),
            PercentualeDocumento(indiceDocumento, totaleDocumenti,
                                indiceRiga + 1, documento.Righe.Count))
            Next

            If Not HaRigheValide() Then
                motivo = "documento privo di righe valide: salvataggio annullato"
                Return False
            End If

            AggiungiNote(noteNonBloccanti)
            MostraAvanzamento("Impegno " & indiceDocumento.ToString() & " di " &
        totaleDocumenti.ToString() & " (" & documento.Riferimento & "): salvataggio...",
        PercentualeDocumento(indiceDocumento, totaleDocumenti,
                                documento.Righe.Count, documento.Righe.Count))
            If Not SalvaImpegno() Then
                motivo = _ultimoErrore
                Return False
            End If

            Return True
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
            motivo = ex.Message
            Return False
        End Try
    End Function

#End Region

#Region "Utils"

    Private Function ConfermaReimportazione(ByVal riferimento As String,
                                            ByVal esistenti As DataTable) As Boolean
        Dim documenti As New List(Of String)()
        For Each riga As DataRow In esistenti.Rows
            Dim serie As String = NTSCStr(riga!td_serie).Trim()
            If serie = "" Then serie = "(vuota)"
            documenti.Add("- Tipo " & NTSCStr(riga!td_tipork).Trim() &
                          ", anno " & NTSCStr(riga!td_anno) &
                          ", serie " & serie &
                          ", numero " & NTSCStr(riga!td_numord) &
                          ", conto " & NTSCStr(riga!td_conto))
        Next

        Dim messaggio As String = "Il riferimento " & riferimento &
                                 " è già presente nei seguenti documenti:" & vbCrLf &
                                 String.Join(vbCrLf, documenti.ToArray()) & vbCrLf & vbCrLf &
                                 "Sei sicuro di voler reimportare questo ordine?"
        Return _oApp.MsgBoxInfoYesNo_DefNo(messaggio) = DialogResult.Yes
    End Function

    Private Sub RisolviDestinazione(ByVal conto As Integer,
                                  ByVal codDestEsterno As Integer,
                                  ByVal dataBase As Date,
                                  ByRef codDest As Integer,
                                  ByRef dataConsegna As Nullable(Of Date),
                                  ByVal noteNonBloccanti As List(Of String))
        Try
            codDest = 0
            Dim risultato As DataTable = _oClfGsor.GetDestinazioneImport(codDestEsterno.ToString(), conto)
            If risultato Is Nothing OrElse risultato.Rows.Count = 0 Then
                noteNonBloccanti.Add("Destinazione esterna " & codDestEsterno.ToString() &
                                     " non trovata: utilizzata la destinazione principale (0).")
                Return
            End If

            codDest = NTSCInt(risultato.Rows(0)!dd_coddest)
            Dim giornoConsegna As String = NTSCStr(risultato.Rows(0)!dd_hhGiornoConsegna).Trim()
            If String.IsNullOrWhiteSpace(giornoConsegna) Then Return

            dataConsegna = CalcolaProssimaDataConsegna(giornoConsegna, dataBase)
            If Not dataConsegna.HasValue Then
                noteNonBloccanti.Add("Giorno consegna non riconosciuto per destinazione " &
                            codDestEsterno.ToString() & ": " & giornoConsegna)
            End If
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
            codDest = 0
            noteNonBloccanti.Add("Errore nella ricerca della destinazione esterna " &
                                 codDestEsterno.ToString() &
                                 ": utilizzata la destinazione principale (0).")
        End Try
    End Sub

    Private Function CalcolaProssimaDataConsegna(ByVal giorno As String,
                                                ByVal dataBase As Date) As Nullable(Of Date)
        Try
            Dim valore As String = NTSCStr(giorno).Trim().ToLowerInvariant().Replace("ì", "i")
            If valore.Length >= 3 Then valore = valore.Substring(0, 3)

            Dim giornoSettimana As DayOfWeek
            Select Case valore
                Case "lun" : giornoSettimana = DayOfWeek.Monday
                Case "mar" : giornoSettimana = DayOfWeek.Tuesday
                Case "mer" : giornoSettimana = DayOfWeek.Wednesday
                Case "gio" : giornoSettimana = DayOfWeek.Thursday
                Case "ven" : giornoSettimana = DayOfWeek.Friday
                Case "sab" : giornoSettimana = DayOfWeek.Saturday
                Case "dom" : giornoSettimana = DayOfWeek.Sunday
                Case Else : Return Nothing
            End Select

            Dim giorniDaAggiungere As Integer =
        (CInt(giornoSettimana) - CInt(dataBase.DayOfWeek) + 7) Mod 7
            If giorniDaAggiungere = 0 Then giorniDaAggiungere = 7
            Return dataBase.Date.AddDays(giorniDaAggiungere)
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
            Return Nothing
        End Try
    End Function

    Private Function PercentualeDocumento(ByVal indiceDocumento As Integer,
                                         ByVal totaleDocumenti As Integer,
                                         ByVal rigaElaborata As Integer,
                                         ByVal totaleRighe As Integer) As Integer
        Try
            If totaleDocumenti <= 0 Then Return 5
            Dim quotaDocumento As Decimal = 90D / totaleDocumenti
            Dim avanzamentoRighe As Decimal = 0D
            If totaleRighe > 0 Then avanzamentoRighe = quotaDocumento * rigaElaborata / totaleRighe
            Return CInt(Math.Truncate(5D + quotaDocumento * (indiceDocumento - 1) + avanzamentoRighe))
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
            Return 0
        End Try
    End Function

    Private Sub MostraAvanzamento(ByVal messaggio As String, ByVal percentuale As Integer)
        Try
            RaiseEvent Avanzamento(messaggio, Math.Max(0, Math.Min(100, percentuale)))
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
        End Try
    End Sub

    Private Function InserisciRiga(ByVal riga As HH_RigaImpegnoImport,
                                  ByVal noteNonBloccanti As List(Of String)) As Boolean
        Try
            'Le righe testuali vengono riportate nelle note senza creare righe articolo.
            If String.Equals(riga.TipoRiga, "T", StringComparison.OrdinalIgnoreCase) Then
                If HaContenutoSignificativo(riga.Descrizione) Then
                    noteNonBloccanti.Add(riga.Descrizione)
                End If
                Return True
            Else
                Dim codiceOriginale As String = riga.Codart
                If Not AggiungiRiga(codiceOriginale, riga.Quantita, "") Then
                    Dim erroreCodart As String = _ultimoErrore
                    noteNonBloccanti.Add(codiceOriginale &
                               " - Quantità: " & NTSCStr(riga.Quantita) &
                               " - " & erroreCodart)
                    Return True
                End If
                AggiungiAvvisoEntity(noteNonBloccanti)
            End If

            If Not SalvaRigaCorrente() Then
                noteNonBloccanti.Add(_ultimoErrore)
            Else
                AggiungiAvvisoEntity(noteNonBloccanti)
            End If

            Return True
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
            Return False
        End Try
    End Function

    Private Function HaContenutoSignificativo(ByVal testo As String) As Boolean
        Try
            If String.IsNullOrWhiteSpace(testo) Then Return False

            'Scarta righe composte esclusivamente da spazi o punteggiatura.
            For Each carattere As Char In testo
                If Char.IsLetterOrDigit(carattere) Then Return True
            Next
            Return False
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
            Return False
        End Try
    End Function

    Private Sub AggiungiAvvisoEntity(ByVal noteNonBloccanti As List(Of String))
        Try
            Dim avviso As String = ConsumaAvvisoEntity()
            If Not String.IsNullOrWhiteSpace(avviso) Then noteNonBloccanti.Add(avviso)
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
        End Try
    End Sub

#End Region

#Region "Lettura e raggruppamento Excel"

    Private Function LeggiERaggruppa(ByVal percorsoFile As String) As List(Of HH_DocumentoImpegnoImport)
        Dim documenti As New List(Of HH_DocumentoImpegnoImport)
        Try
            Dim tabella As DataTable = LeggiExcel(percorsoFile)
            Dim perRiferimento As New Dictionary(Of String, HH_DocumentoImpegnoImport)(StringComparer.OrdinalIgnoreCase)

            For Each rigaExcel As DataRow In tabella.Rows
                Dim riferimento As String = ValoreTesto(rigaExcel, 2).Trim()       'Colonna C
                If String.IsNullOrWhiteSpace(riferimento) OrElse riferimento = "bc_id" Then Continue For

                Dim documento As HH_DocumentoImpegnoImport = Nothing
                If Not perRiferimento.TryGetValue(riferimento, documento) Then
                    documento = New HH_DocumentoImpegnoImport() With {
            .Riferimento = riferimento,
            .Cliente = ValoreTesto(rigaExcel, 0).Trim(),                  'Colonna A
            .CodDest = NTSCInt(ValoreTesto(rigaExcel, 5)),               'Colonna F
            .DataBaseConsegna = LeggiDataBaseConsegna(rigaExcel(7), riferimento) 'Colonna H
          }
                    perRiferimento.Add(riferimento, documento)
                    documenti.Add(documento)
                End If

                documento.Righe.Add(New HH_RigaImpegnoImport() With {
          .TipoRiga = ValoreTesto(rigaExcel, 8).Trim(),                   'Colonna I
          .Descrizione = ValoreTesto(rigaExcel, 11).Trim(),              'Colonna L
          .Quantita = NTSCDec(ValoreTesto(rigaExcel, 12)),               'Colonna M
          .Codart = CreaCodarfo(rigaExcel)                               'Colonne J-K
        })
            Next

        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
            Throw
        End Try
        Return documenti
    End Function

    Private Function CreaCodarfo(ByVal riga As DataRow) As String
        Try
            Dim prefisso As String = ValoreTesto(riga, 9).Trim()                'Colonna J
            Dim codice As String = ValoreTesto(riga, 10).Trim()                'Colonna K

            If String.IsNullOrWhiteSpace(prefisso) Then Return codice
            If String.IsNullOrWhiteSpace(codice) Then Return prefisso
            Return prefisso.TrimEnd("-"c) & "-" & codice.TrimStart("-"c)
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
            Return ""
        End Try
    End Function

    Private Function LeggiExcel(ByVal percorsoFile As String) As DataTable
        Dim risultato As New DataTable()
        Try
            Using flusso As FileStream = File.Open(percorsoFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                Using lettore As IExcelDataReader = ExcelReaderFactory.CreateReader(flusso)
                    Dim foglioTrovato As Boolean = False
                    Do
                        If String.Equals(lettore.Name, "faxb2b", StringComparison.OrdinalIgnoreCase) Then
                            foglioTrovato = True
                            Exit Do
                        End If
                    Loop While lettore.NextResult()

                    If Not foglioTrovato Then
                        Throw New Exception("Il file Excel non contiene il foglio richiesto 'faxb2b'.")
                    End If

                    For indiceColonna As Integer = 0 To lettore.FieldCount - 1
                        risultato.Columns.Add("F" & indiceColonna.ToString(), GetType(Object))
                    Next

                    While lettore.Read()
                        Dim riga As DataRow = risultato.NewRow()
                        For indiceColonna As Integer = 0 To lettore.FieldCount - 1
                            Dim valore As Object = lettore.GetValue(indiceColonna)
                            riga(indiceColonna) = If(valore Is Nothing, DBNull.Value, valore)
                        Next
                        risultato.Rows.Add(riga)
                    End While
                End Using
            End Using

            If risultato.Columns.Count < 15 Then
                Throw New Exception("Il file Excel non contiene tutte le colonne richieste (A-O).")
            End If
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
            Throw
        End Try
        Return risultato
    End Function

    Private Function LeggiDataBaseConsegna(ByVal valore As Object,
                                           ByVal riferimento As String) As Date
        If TypeOf valore Is DateTime Then Return CType(valore, DateTime).Date
        If TypeOf valore Is Double Then Return DateTime.FromOADate(CDbl(valore)).Date

        Dim data As DateTime
        If valore IsNot Nothing AndAlso valore IsNot DBNull.Value AndAlso
           DateTime.TryParse(NTSCStr(valore), CultureInfo.GetCultureInfo("it-IT"),
                             DateTimeStyles.None, data) Then Return data.Date

        Throw New Exception("Data non valida nella colonna H per l'impegno " & riferimento & ".")
    End Function

    Private Function ValoreTesto(ByVal riga As DataRow, ByVal indice As Integer) As String
        Try
            If riga Is Nothing OrElse indice < 0 OrElse indice >= riga.ItemArray.Length Then Return ""
            If riga(indice) Is DBNull.Value Then Return ""
            Return NTSCStr(riga(indice))
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
            Return ""
        End Try
    End Function

#End Region

#Region "Creazione impegno"

    Private Function InizializzaEntityOrdini() As Boolean
        Try
            PulisciErrore()
            Dim strErr As String = ""
            Dim oTmp As Object = Nothing
            If Not CLN__STD.NTSIstanziaDll(_oApp.ServerDir, _oApp.NetDir, "BNORGSOR", "BEORGSOR",
                                    oTmp, strErr, False, "", "") Then
                _ultimoErrore = "inizializzazione BEORGSOR non riuscita: " & strErr
                Return False
            End If

            _oCleGsor = CType(oTmp, CLEORGSOR)
            AddHandler _oCleGsor.RemoteEvent, AddressOf GestisciEventoEntity
            If Not _oCleGsor.Init(_oApp, Nothing, _oMenu.oCleComm, "", False, "", "") Then
                _ultimoErrore = "inizializzazione ordine non riuscita"
                Return False
            End If
            If Not _oCleGsor.InitExt() Then
                _ultimoErrore = "inizializzazione estesa ordine non riuscita"
                Return False
            End If
            Return True
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
            _ultimoErrore = ex.Message
            Return False
        End Try
    End Function

    Private Function NuovoImpegno(ByVal conto As Integer, ByVal riferimento As String,
                          ByVal codDest As Integer,
                          ByVal dataConsegna As Nullable(Of Date)) As Boolean
        Try
            PulisciErrore()
            Dim tipork As String = "R"
            Dim serie As String = " "
            Dim anno As Integer = Now.Year
            Dim numero As Integer = _oCleGsor.LegNuma(tipork, serie, anno)
            If numero = 0 Then
                _ultimoErrore = "numerazione impegni R non disponibile"
                Return False
            End If

            ImpostaFlagCreazione()
            If Not _oCleGsor.ApriOrdine(_oApp.Ditta, True, tipork, anno, serie, numero,
                                  _oCleGsor.dsShared) Then
                _ultimoErrore = MotivoEntity("apertura preventiva impegno non riuscita")
                Return False
            End If
            If _oCleGsor.dsShared.Tables.Contains("TESTA") AndAlso
         _oCleGsor.dsShared.Tables("TESTA").Rows.Count > 0 Then
                _ultimoErrore = "il numero assegnato risulta già esistente"
                Return False
            End If

            _oCleGsor.bInApriDocSilent = True
            _oCleGsor.ResetVar()
            _oCleGsor.strVisNoteConto = "N"
            ImpostaFlagCreazione()
            _oCleGsor.NuovoOrdine(_oApp.Ditta, tipork, anno, serie, numero, "")
            _oCleGsor.bInNuovoDocSilent = True

            If _oCleGsor.dttET Is Nothing OrElse _oCleGsor.dttET.Rows.Count = 0 Then
                _ultimoErrore = MotivoEntity("testata impegno non inizializzata")
                Return False
            End If

            With _oCleGsor.dttET.Rows(0)
                !codditt = _oApp.Ditta
                !et_conto = conto
                !et_tipobf = 1
                !et_scopag = 0
                !et_riferim = riferimento
                !et_coddest = codDest
                If dataConsegna.HasValue Then !et_datcons = dataConsegna.Value
            End With
            Return True
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
            _ultimoErrore = ex.Message
            Return False
        End Try
    End Function

    Private Function AggiungiRiga(ByVal codart As String, ByVal quantita As Decimal,
                               ByVal descrizione As String) As Boolean
        Try
            PulisciErrore()
            Dim numeroRigheIniziale As Integer = 0
            If _oCleGsor.dttEC IsNot Nothing Then numeroRigheIniziale = _oCleGsor.dttEC.Rows.Count

            Dim numeroRiga As Integer = GetUltimoNumeroRiga() + 1
            If Not _oCleGsor.AggiungiRigaCorpo(False, codart, 0, numeroRiga) Then
                _ultimoErrore = MotivoEntity("articolo non riconosciuto: " & codart)
                RimuoviRigheAggiunte(numeroRigheIniziale)
                Return False
            End If

            'Business può restituire una riga formalmente aggiunta ma priva del codice articolo.
            If _oCleGsor.dttEC Is Nothing OrElse
         _oCleGsor.dttEC.Rows.Count <= numeroRigheIniziale OrElse
         String.IsNullOrWhiteSpace(NTSCStr(_oCleGsor.dttEC.Rows(_oCleGsor.dttEC.Rows.Count - 1)!ec_codart)) Then
                _ultimoErrore = MotivoEntity("articolo non riconosciuto: " & codart)
                RimuoviRigheAggiunte(numeroRigheIniziale)
                Return False
            End If

            Dim riga As DataRow = _oCleGsor.dttEC.Rows(_oCleGsor.dttEC.Rows.Count - 1)
            riga!ec_quant = quantita
            If Not String.IsNullOrWhiteSpace(descrizione) Then riga!ec_descr = descrizione
            Return True
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
            _ultimoErrore = ex.Message
            Return False
        End Try
    End Function

    Private Function SalvaRigaCorrente() As Boolean
        Try
            PulisciErrore()
            If _oCleGsor.dttEC Is Nothing OrElse _oCleGsor.dttEC.Rows.Count = 0 Then
                _ultimoErrore = "nessuna riga corpo disponibile"
                Return False
            End If
            If Not _oCleGsor.RecordSalva(_oCleGsor.dttEC.Rows.Count - 1, False, Nothing) Then
                _ultimoErrore = MotivoEntity("salvataggio riga non riuscito")
                Return False
            End If
            Return True
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
            _ultimoErrore = ex.Message
            Return False
        End Try
    End Function

    Private Function HaRigheValide() As Boolean
        Try
            If _oCleGsor Is Nothing OrElse _oCleGsor.dttEC Is Nothing OrElse
               Not _oCleGsor.dttEC.Columns.Contains("ec_codart") Then Return False

            For Each riga As DataRow In _oCleGsor.dttEC.Rows
                If riga.RowState <> DataRowState.Deleted AndAlso
                   Not String.IsNullOrWhiteSpace(NTSCStr(riga!ec_codart)) Then Return True
            Next
            Return False
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
            Return False
        End Try
    End Function

    Private Sub AggiungiNote(ByVal note As List(Of String))
        Try
            If note Is Nothing OrElse note.Count = 0 Then Return
            Dim testoEsistente As String = NTSCStr(_oCleGsor.dttET.Rows(0)!et_note).Trim()
            Dim testoNuovo As String = String.Join(vbCrLf, note.ToArray())
            _oCleGsor.dttET.Rows(0)!et_note =
        If(String.IsNullOrWhiteSpace(testoEsistente), testoNuovo,
           testoEsistente & vbCrLf & testoNuovo)
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
        End Try
    End Sub

    Private Function SalvaImpegno() As Boolean
        Try
            PulisciErrore()
            If Not _oCleGsor.OkTestata() Then
                _ultimoErrore = MotivoEntity("controlli testata non superati")
                Return False
            End If
            _oCleGsor.CalcolaTotali()
            If Not _oCleGsor.SalvaOrdine("N") Then
                _ultimoErrore = MotivoEntity("salvataggio impegno non riuscito")
                Return False
            End If
            Return True
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
            _ultimoErrore = ex.Message
            Return False
        End Try
    End Function

    Private Function ConsumaAvvisoEntity() As String
        Try
            Dim messaggio As String = _ultimoMessaggioEntity
            _ultimoMessaggioEntity = ""
            Return messaggio
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
            Return ""
        End Try
    End Function

#End Region

#Region "Utils_ORGSOR"

    Private Sub RimuoviRigheAggiunte(ByVal numeroRigheIniziale As Integer)
        Try
            If _oCleGsor Is Nothing OrElse _oCleGsor.dttEC Is Nothing Then Return

            'Ripristina il corpo allo stato precedente al tentativo non riuscito.
            While _oCleGsor.dttEC.Rows.Count > numeroRigheIniziale
                _oCleGsor.dttEC.Rows.RemoveAt(_oCleGsor.dttEC.Rows.Count - 1)
            End While
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
        End Try
    End Sub

    Private Sub ImpostaFlagCreazione()
        Try
            _oCleGsor.bDisabilitaCheckAnnoData = True
            _oCleGsor.bDisabilitaCheckDateAnteriori = True
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
        End Try
    End Sub

    Private Function GetUltimoNumeroRiga() As Integer
        Try
            If _oCleGsor.dttEC Is Nothing OrElse _oCleGsor.dttEC.Rows.Count = 0 Then Return 0
            Return NTSCInt(_oCleGsor.dttEC.Rows(_oCleGsor.dttEC.Rows.Count - 1)!ec_riga)
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
            Return 0
        End Try
    End Function

    Private Sub GestisciEventoEntity(ByVal sender As Object, ByRef e As NTSEventArgs)
        Try
            If e IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(e.Message) Then
                _ultimoMessaggioEntity = e.Message.Replace(vbCr, " ").Replace(vbLf, " ").Trim()
            End If
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
        End Try
    End Sub

    Private Sub PulisciErrore()
        Try
            _ultimoErrore = ""
            _ultimoMessaggioEntity = ""
    Catch ex As Exception
      CLN__STD.GestErr(ex, Me, "")
    End Try
  End Sub

    Private Function MotivoEntity(ByVal fallback As String) As String
        Try
            If Not String.IsNullOrWhiteSpace(_ultimoMessaggioEntity) Then Return _ultimoMessaggioEntity
            Return fallback
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
            Return fallback
        End Try
    End Function

#End Region

End Class
