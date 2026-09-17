Imports System.Data
Imports System.IO
Imports ExcelDataReader
Imports System.Globalization
Imports NTSInformatica.CLN__STD
Imports System
Public Class FROORGSOR
    Inherits FRMORGSOR

    Public _oClfGsor As CLFORGSOR

    Public WithEvents edet_hhtotvol As NTSInformatica.NTSTextBoxNum
    Public WithEvents lbet_hhtotvol As NTSInformatica.NTSLabel
    Public WithEvents ec_hhvol As NTSInformatica.NTSGridColumn
    Public WithEvents ec_hhisetrfid As NTSInformatica.NTSGridColumn
    Public WithEvents edet_hhtotpn As NTSInformatica.NTSTextBoxNum
    Public WithEvents lbet_hhtotpn As NTSInformatica.NTSLabel
    Public WithEvents ec_hhpn As NTSInformatica.NTSGridColumn
    Public WithEvents edet_hhtotpl As NTSInformatica.NTSTextBoxNum
    Public WithEvents lbet_hhtotpl As NTSInformatica.NTSLabel
    Public WithEvents ec_hhpl As NTSInformatica.NTSGridColumn
    Public WithEvents fm_hhPostiLettoCalcolati As NTSInformatica.NTSGroupBox
    Public WithEvents lb_hhPostiLettoCalcolati As NTSInformatica.NTSLabel
    Public WithEvents ed_hhPostiLettoCalcolati As NTSInformatica.NTSTextBoxNum
    Public WithEvents tlb_hhImportaImpegni As NTSInformatica.NTSBarButtonItem
    Private Shared strUltimaCartellaImport As String = ""


    Public Property oClfGsor() As CLFORGSOR
        Get
            If _oClfGsor Is Nothing Then _oClfGsor = CType(oCleGsor, CLFORGSOR)
            Return _oClfGsor
        End Get
        Set(ByVal value As CLFORGSOR)
            _oClfGsor = value
        End Set
    End Property
    Public Overrides Sub InitializeComponent()
        MyBase.InitializeComponent()
        Me.fm_hhPostiLettoCalcolati = New NTSInformatica.NTSGroupBox()
        Me.lb_hhPostiLettoCalcolati = New NTSInformatica.NTSLabel()
        Me.ed_hhPostiLettoCalcolati = New NTSInformatica.NTSTextBoxNum()
        Me.fm_hhPostiLettoCalcolati.Name = "fm_hhPostiLettoCalcolati"
        Me.fm_hhPostiLettoCalcolati.Text = "POSTI LETTO"
        Me.fm_hhPostiLettoCalcolati.Tile = True
        Me.fm_hhPostiLettoCalcolati.TileIndex = 1
        Me.fm_hhPostiLettoCalcolati.Margin = New System.Windows.Forms.Padding(4)
        Me.fm_hhPostiLettoCalcolati.Size = New System.Drawing.Size(200, 80)
        Me.lb_hhPostiLettoCalcolati.Name = "lb_hhPostiLettoCalcolati"
        Me.lb_hhPostiLettoCalcolati.Text = "Posti Letto Calcolati"
        Me.lb_hhPostiLettoCalcolati.BackColor = System.Drawing.Color.Transparent
        Me.lb_hhPostiLettoCalcolati.Location = New System.Drawing.Point(4, 24)
        Me.lb_hhPostiLettoCalcolati.Size = New System.Drawing.Size(188, 20)
        Me.lb_hhPostiLettoCalcolati.NTSBordeStyle = NTSInformatica.NTSLabel.NTSBorderStyle.FieldCaption
        Me.ed_hhPostiLettoCalcolati.Name = "ed_hhPostiLettoCalcolati"
        Me.ed_hhPostiLettoCalcolati.Location = New System.Drawing.Point(4, 48)
        Me.ed_hhPostiLettoCalcolati.Size = New System.Drawing.Size(88, 20)
        Me.ed_hhPostiLettoCalcolati.Properties.AutoHeight = False
        Me.fm_hhPostiLettoCalcolati.Controls.Add(Me.lb_hhPostiLettoCalcolati)
        Me.fm_hhPostiLettoCalcolati.Controls.Add(Me.ed_hhPostiLettoCalcolati)
        Me.flAltro.Controls.Add(Me.fm_hhPostiLettoCalcolati)
        Me.flAltro.Controls.SetChildIndex(Me.fm_hhPostiLettoCalcolati, 1)
        Me.tlb_hhImportaImpegni = New NTSInformatica.NTSBarButtonItem()
        Me.NtsBarManager1.Items.AddRange(New DevExpress.XtraBars.BarItem() {Me.tlb_hhImportaImpegni})
        Me.NtsBarManager1.MaxItemId = 214
        Me.tlbMain.LinksPersistInfo.AddRange(New DevExpress.XtraBars.LinkPersistInfo() {
New DevExpress.XtraBars.LinkPersistInfo(Me.tlb_hhImportaImpegni)})
        '
        'tlb_hhImportaImpegni
        '
        Me.tlb_hhImportaImpegni.Caption = "Importa impegni"
        Me.tlb_hhImportaImpegni.Glyph = GetGlyphImportaImpegni()
        Me.tlb_hhImportaImpegni.Id = 213
        Me.tlb_hhImportaImpegni.Name = "tlb_hhImportaImpegni"
        Me.tlb_hhImportaImpegni.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionInMenu
        Me.tlb_hhImportaImpegni.Visible = False
    End Sub
    Public Overrides Sub InitControls()
        MyBase.InitControls()
        Try
            ed_hhPostiLettoCalcolati.NTSSetParam(oMenu, "Posti Letto Calcolati", "0.00")
            edet_hhtotvol = CType(NTSFindControlByName(Me, "edet_hhtotvol"), NTSTextBoxNum)
            lbet_hhtotvol = CType(NTSFindControlByName(Me, "lbet_hhtotvol"), NTSLabel)
            ec_hhvol = grvRighe.NTSGetColumnByName("ec_hhvol")
            edet_hhtotvol.NTSSetParam(oMenu, "Volume", oApp.FormatQta)
            ec_hhvol.NTSSetParamNUM(oMenu, "Volume per riga", oApp.FormatQta, 13)
            ec_hhisetrfid = grvRighe.NTSGetColumnByName("ec_hhisetrfid")
            edet_hhtotpn = CType(NTSFindControlByName(Me, "edet_hhtotpn"), NTSTextBoxNum)
            lbet_hhtotpn = CType(NTSFindControlByName(Me, "lbet_hhtotpn"), NTSLabel)
            ec_hhpn = grvRighe.NTSGetColumnByName("ec_hhpn")
            edet_hhtotpn.NTSSetParam(oMenu, "Peso netto", oApp.FormatQta)
            ec_hhpn.NTSSetParamNUM(oMenu, "Peso netto per riga", oApp.FormatQta, 13)
            edet_hhtotpl = CType(NTSFindControlByName(Me, "edet_hhtotpl"), NTSTextBoxNum)
            lbet_hhtotpl = CType(NTSFindControlByName(Me, "lbet_hhtotpl"), NTSLabel)
            ec_hhpl = grvRighe.NTSGetColumnByName("ec_hhpl")
            edet_hhtotpl.NTSSetParam(oMenu, "Peso lordo", oApp.FormatQta)
            ec_hhpl.NTSSetParamNUM(oMenu, "Peso lordo per riga", oApp.FormatQta, 13)
        Catch ex As Exception
            '-------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '-------------------------------------------------
        End Try
    End Sub
    Public Overrides Sub Bindcontrols()
        MyBase.Bindcontrols()
        ed_hhPostiLettoCalcolati.NTSDbField = "TESTA.et_hhPostiLettoCalcolati"
        NTSFormAddDataBinding(dcGsor, fm_hhPostiLettoCalcolati)
    End Sub
    Public Overrides Sub FRMORGSOR_Load(sender As Object, e As EventArgs)
        Try
            MyBase.FRMORGSOR_Load(sender, e)
            ed_hhPostiLettoCalcolati.Enabled = False
            edet_hhtotvol.Enabled = False
            ec_hhvol.Enabled = False
            'ec_hhisetrfid.Visible = CBool(IIf(cbTipoDoc.SelectedValue.ToString <> "H", False, True))
            lbet_hhtotvol.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            edet_hhtotpn.Enabled = False
            ec_hhpn.Enabled = False
            lbet_hhtotpn.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            edet_hhtotpl.Enabled = False
            ec_hhpl.Enabled = False
            lbet_hhtotpl.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
        End Try
    End Sub
    Public Overrides Sub cbTipoDoc_SelectedIndexChanged(sender As Object, e As EventArgs)
        Try

            MyBase.cbTipoDoc_SelectedIndexChanged(sender, e)
            'Se Ordine di Produzione abilito, in griglia, il campo/colonna pers. che indica se � una riga che dovr� andare nelle etichette rfid
            If Not IsNothing(ec_hhisetrfid) Then
                If cbTipoDoc.SelectedValue.ToString <> "H" Then
                    ec_hhisetrfid.Enabled = False
                Else
                    ec_hhisetrfid.Enabled = True
                End If
            End If
            'Gestione visibilità pulsante di importazione impegni.
            If cbTipoDoc.Text = "Impegno cliente" Then
                Me.tlb_hhImportaImpegni.Visible = True
            Else
                Me.tlb_hhImportaImpegni.Visible = False
            End If

        Catch ex As Exception
            '-------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '-------------------------------------------------	
        End Try
    End Sub
    Public Overridable Sub edNumDoc_Validated(sender As Object, e As EventArgs) Handles edNumDoc.Validated
        Try

            If edNumDoc.TextInt = 0 Then
                Me.tlb_hhImportaImpegni.Visible = True
            Else
                Me.tlb_hhImportaImpegni.Visible = False
            End If

        Catch ex As Exception
            '-------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '-------------------------------------------------	
        End Try

    End Sub
    'Public Overrides Function Salva(nStampa As Integer, bAllowCancel As Boolean) As Boolean
    '  Try
    '    'Lancio prima il calcolo dei totali
    '    oCleGsor.dttET.Rows(0)!et_hhtotvol = oClfGsor.CalcolaTotaleVolume()
    '    Return MyBase.Salva(nStampa, bAllowCancel)
    '  Catch ex As Exception
    '    Dim strErr As String = GestError(ex, Me, "", oApp.InfoError, oApp.ErrorLogFile, True)
    '  End Try
    'End Function


    Public Overridable Sub tlb_hhImportaImpegni_ItemClick(ByVal sender As Object,
                                                          ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles tlb_hhImportaImpegni.ItemClick
        Dim frmProgress As BOORGSOR_ProgressView = Nothing

        Try
            Dim percorsoFile As String = SelezionaFileImpegni()
            If String.IsNullOrWhiteSpace(percorsoFile) Then Return

            tlb_hhImportaImpegni.Enabled = False

            Try
                frmProgress = New BOORGSOR_ProgressView()
                If Not frmProgress.Init(CType(oApp.oMenu, CLE__MENU), Nothing, oApp.Ditta) Then Return

                frmProgress.Show(Me)
                frmProgress.Aggiorna("Avvio importazione impegni...", 0)

                Dim esito As EsitoImportazioneImpegni = ImportaImpegni(percorsoFile, frmProgress)
                MostraEsitoImportazione(esito)

            Finally
                If frmProgress IsNot Nothing Then frmProgress.Close()
                tlb_hhImportaImpegni.Enabled = True
            End Try

        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
        End Try
    End Sub


#Region "ORGSOR"

    Private Function GsorInit() As Boolean
        Try
            If oClfGsor Is Nothing Then Return False

            If dsGsor Is Nothing OrElse dsGsor.Tables.Count < 3 Then
                If Not oClfGsor.ApriOrdine(DittaCorrente, True, "R", 0, " ", 0, dsGsor) Then
                    Return False
                End If
            End If

            Return True

        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
            Return False
        End Try
    End Function

    Private Function GsorCrea(ByVal nConto As Integer,
                              ByVal nDestinazione As Integer,
                              ByVal strRiferimento As String,
                              ByVal dataConsegna As Nullable(Of Date)) As Boolean
        Try

            cbTipoDoc.SelectedValue = "R"

            If edSerieDoc.Text = "" Then edSerieDoc.Text = " "
            edSerieDoc.Text = edSerieDoc.Text.ToUpper()

            Dim nAnno As Integer = Now.Year
            edAnnoDoc.Text = nAnno.ToString()

            Dim nNumero As Integer = oClfGsor.LegNuma("R", edSerieDoc.Text, nAnno)
            If nNumero = 0 Then Return False

            edNumDoc.Text = nNumero.ToString()

            bInLoadOrd = True

            If dsGsor Is Nothing OrElse dsGsor.Tables.Count < 3 Then
                If Not oClfGsor.ApriOrdine(
                    DittaCorrente, True, "R", 0, " ", 0, dsGsor) Then
                    bInLoadOrd = False
                    Return False
                End If
            End If

            If Not oClfGsor.NuovoOrdine(
                DittaCorrente, "R", nAnno, edSerieDoc.Text, nNumero,
                ceModd.strModelloCodice) Then
                bInLoadOrd = False
                Return False
            End If

            'Lo standard usa questi flag nei flussi di generazione automatica.
            'bInNuovoDocSilent evita le richieste YES/NO durante il salvataggio;
            'bInCreaDocDaGnor evita ulteriori messaggi tipici della creazione automatica.
            oClfGsor.bInNuovoDocSilent = True
            oClfGsor.bInCreaDocDaGnor = True
            oClfGsor.bDisabilitaMsgPrezzoZero = True
            oClfGsor.bDisabilitaMsgQuantZero = True
            oClfGsor.bDisabilitaCheckAnnoData = True
            oClfGsor.bDisabilitaCheckDateAnteriori = True
            oClfGsor.strVisNoteConto = "N"
            'Evita l'apertura di FRMMGDATC durante SalvaOrdine.
            oClfGsor.bVerificaDateConsegnaIC = False

            dsGsor.Tables("TESTA").AcceptChanges()

            dcGsor.DataSource = dsGsor.Tables("TESTA")
            Bindcontrols()

            If dsGsor.Tables.Contains("ALLOLET") AndAlso
               dsGsor.Tables("ALLOLET").Rows.Count > 0 Then
                dsGsor.Tables("ALLOLET").Clear()
                dsGsor.Tables("ALLOLET").AcceptChanges()
            End If

            dcGsorRighe.DataSource = dsGsor.Tables("CORPO")
            grRighe.DataSource = dcGsorRighe

            bInLoadOrd = False

            SetUnboundFiled(0)
            NumeroPost()

            'Conto tramite il controllo della BNORGSOR, così restano attive
            'le normali proposte Business.
            edEt_conto.Text = NTSCStr(nConto)
            edEt_conto.NTSTextDB = NTSCStr(nConto)
            Me.ValidaLastControl()

            If NTSCInt(dsGsor.Tables("TESTA").Rows(0)!et_conto) <> nConto Then
                Return False
            End If

            'Destinazione sempre dopo il conto.
            If nDestinazione <> 0 Then
                edEt_coddest.NTSTextDB = NTSCStr(nDestinazione)
                Me.ValidaLastControl()

                If NTSCInt(dsGsor.Tables("TESTA").Rows(0)!et_coddest) <> nDestinazione Then
                    Return False
                End If
            End If

            'Campi dell'import già presenti nella precedente implementazione.
            With dsGsor.Tables("TESTA").Rows(0)
                !et_riferim = strRiferimento
                !et_hhulrif = strRiferimento
                !et_idtracking = strRiferimento
                If dataConsegna.HasValue Then !et_datcons = dataConsegna.Value
            End With

            Return True

        Catch ex As Exception
            bInLoadOrd = False
            CLN__STD.GestErr(ex, Me, "")
            Return False
        End Try
    End Function

    Private Function GsorAggiungiRiga(ByVal strCodArt As String,
                                      ByVal nQuantita As Decimal) As Boolean
        Try
            If dsGsor Is Nothing OrElse
               Not dsGsor.Tables.Contains("CORPO") OrElse
               Not dsGsor.Tables.Contains("TESTA") Then
                Return False
            End If

            Dim dttEC As DataTable = dsGsor.Tables("CORPO")
            Dim dtrET As DataRow = dsGsor.Tables("TESTA").Rows(0)
            Dim nRighePrima As Integer = dttEC.Rows.Count

            Dim dtrEC As DataRow = dttEC.NewRow()
            dttEC.Rows.Add(dtrEC)

            With dtrEC
                'Forza MovordOnAddNewRow come nello standard BNORGSOR.
                !codditt = "."
                !codditt = DittaCorrente

                'I magazzini vanno impostati prima dell'articolo.
                !ec_magaz = dtrET!et_magaz
                !ec_magaz2 = dtrET!et_magaz2

                !ec_codart = strCodArt
                !ec_quant = nQuantita
                !ec_colli = nQuantita
            End With

            If String.IsNullOrWhiteSpace(NTSCStr(dtrEC!ec_codart)) Then
                If dttEC.Rows.Count > nRighePrima Then dttEC.Rows.RemoveAt(dttEC.Rows.Count - 1)
                Return False
            End If

            If Not oClfGsor.RecordSalva(dttEC.Rows.IndexOf(dtrEC), False, Nothing) Then
                If dttEC.Rows.Count > nRighePrima Then dttEC.Rows.RemoveAt(dttEC.Rows.Count - 1)
                Return False
            End If

            Return True

        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
            Return False
        End Try
    End Function

    Private Sub GsorAggiungiNota(ByVal testo As String)
        Try
            If String.IsNullOrWhiteSpace(testo) Then Return
            If dsGsor Is Nothing OrElse
               Not dsGsor.Tables.Contains("TESTA") OrElse
               dsGsor.Tables("TESTA").Rows.Count = 0 Then Return

            Dim rigaTesta As DataRow = dsGsor.Tables("TESTA").Rows(0)
            Dim testoEsistente As String = NTSCStr(rigaTesta!et_testonota).Trim()

            rigaTesta!et_testonota =
                If(String.IsNullOrWhiteSpace(testoEsistente),
                   testo.Trim(),
                   testoEsistente & vbCrLf & testo.Trim())

        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
        End Try
    End Sub

    Private Function GsorSalva() As Boolean
        Try
            If Not oClfGsor.OkTestata() Then Return False
            oClfGsor.CalcolaTotali()
            Return oClfGsor.SalvaOrdine("N")

        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
            Return False
        End Try
    End Function

#End Region

#Region "IMPORT EXCEL"

    Private Class RigaImpegnoImport
        Public Property TipoRiga As String
        Public Property CodArt As String
        Public Property Descrizione As String
        Public Property Quantita As Decimal
    End Class

    Private Class DocumentoImpegnoImport
        Public Property Riferimento As String
        Public Property Cliente As String
        Public Property CodDestEsterno As Integer
        Public Property DataBaseConsegna As Date
        Public Property ConsegnaPrecedente As Boolean
        Public ReadOnly Property Righe As New List(Of RigaImpegnoImport)
    End Class

    Private Class EsitoImportazioneImpegni
        Public Property DocumentiSalvati As Integer
        Public Property DocumentiSaltati As Integer
        Public ReadOnly Property ErroriDocumento As New List(Of String)
    End Class

    Private ReadOnly hhClienti As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase) From {
        {"Iris Mobili S.r.l.", 2020018}
    }

    Private Function ImportaImpegni(ByVal percorsoFile As String,
                                    ByVal frmProgress As BOORGSOR_ProgressView) As EsitoImportazioneImpegni
        Dim esito As New EsitoImportazioneImpegni()

        Try
            frmProgress.Aggiorna("Lettura del file Excel...", 0)

            Dim documenti As List(Of DocumentoImpegnoImport) = LeggiERaggruppa(percorsoFile)

            frmProgress.Aggiorna("File letto: " & documenti.Count.ToString() &
                                  " impegni da elaborare.", 5)

            If Not GsorInit() Then
                esito.ErroriDocumento.Add("Inizializzazione ORGSOR non riuscita.")
                Return esito
            End If

            For i As Integer = 0 To documenti.Count - 1
                Dim documento As DocumentoImpegnoImport = documenti(i)
                Dim nDoc As Integer = i + 1

                Try
                    'Mantengo il controllo di reimportazione della versione precedente.
                    Dim esistenti As DataTable =
                        oClfGsor.GetOrdiniPerRiferimento(documento.Riferimento)

                    If esistenti Is Nothing Then
                        Throw New Exception("verifica degli ordini esistenti non riuscita")
                    End If

                    If esistenti.Rows.Count > 0 AndAlso
                       Not ConfermaReimportazione(documento.Riferimento, esistenti) Then
                        esito.DocumentiSaltati += 1
                        Continue For
                    End If

                    frmProgress.Aggiorna(
                        "Impegno " & nDoc.ToString() & " di " & documenti.Count.ToString() &
                        " (" & documento.Riferimento & "): creazione...",
                        PercentualeDocumento(nDoc, documenti.Count, 0, documento.Righe.Count))

                    Dim conto As Integer = 0
                    If Not hhClienti.TryGetValue(documento.Cliente.Trim(), conto) Then
                        Throw New Exception("cliente non configurato: " & documento.Cliente)
                    End If

                    Dim noteNonBloccanti As New List(Of String)()
                    Dim codDest As Integer = 0
                    Dim dataConsegna As Nullable(Of Date) = Nothing

                    RisolviDestinazione(
                        conto,
                        documento.CodDestEsterno,
                        documento.DataBaseConsegna,
                        documento.ConsegnaPrecedente,
                        codDest,
                        dataConsegna,
                        noteNonBloccanti)

                    If Not GsorCrea(
                        conto,
                        codDest,
                        documento.Riferimento,
                        dataConsegna) Then
                        Throw New Exception("creazione impegno non riuscita")
                    End If

                    Dim righeInserite As Integer = 0

                    For r As Integer = 0 To documento.Righe.Count - 1
                        Dim riga As RigaImpegnoImport = documento.Righe(r)

                        'Le righe I="T" sono righe testuali e non generano corpo.
                        If Not String.Equals(riga.TipoRiga, "T", StringComparison.OrdinalIgnoreCase) Then
                            If GsorAggiungiRiga(riga.CodArt, riga.Quantita) Then
                                righeInserite += 1
                            Else
                                noteNonBloccanti.Add(
                                    riga.CodArt &
                                    " - Quantità: " & NTSCStr(riga.Quantita) &
                                    " - articolo non riconosciuto/non inserito")
                            End If
                        End If

                        frmProgress.Aggiorna(
                            "Impegno " & nDoc.ToString() & " di " & documenti.Count.ToString() &
                            " (" & documento.Riferimento & "): riga " &
                            (r + 1).ToString() & " di " & documento.Righe.Count.ToString(),
                            PercentualeDocumento(nDoc, documenti.Count, r + 1, documento.Righe.Count))
                    Next

                    If righeInserite = 0 Then
                        Throw New Exception("documento privo di righe valide")
                    End If

                    For Each nota As String In noteNonBloccanti
                        GsorAggiungiNota(nota)
                    Next

                    frmProgress.Aggiorna(
                        "Impegno " & nDoc.ToString() & " di " & documenti.Count.ToString() &
                        " (" & documento.Riferimento & "): salvataggio...",
                        PercentualeDocumento(nDoc, documenti.Count,
                                             documento.Righe.Count, documento.Righe.Count))

                    If Not GsorSalva() Then
                        Throw New Exception("salvataggio impegno non riuscito")
                    End If

                    esito.DocumentiSalvati += 1

                Catch exDoc As Exception
                    esito.ErroriDocumento.Add(documento.Riferimento & ": " & exDoc.Message)
                End Try
            Next

            frmProgress.Aggiorna("Importazione impegni completata.", 100)

        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
            esito.ErroriDocumento.Add("File: " & ex.Message)
        End Try

        Return esito
    End Function

    Private Function LeggiERaggruppa(ByVal percorsoFile As String) As List(Of DocumentoImpegnoImport)
        Dim documenti As New List(Of DocumentoImpegnoImport)()

        Dim tabella As DataTable = LeggiExcel(percorsoFile)
        Dim perRiferimento As New Dictionary(Of String, DocumentoImpegnoImport)(StringComparer.OrdinalIgnoreCase)

        For Each rigaExcel As DataRow In tabella.Rows
            Dim riferimentoBase As String = ValoreTesto(rigaExcel, 2).Trim()       'C
            If String.IsNullOrWhiteSpace(riferimentoBase) OrElse
               String.Equals(riferimentoBase, "bc_id", StringComparison.OrdinalIgnoreCase) Then
                Continue For
            End If

            Dim documento As DocumentoImpegnoImport = Nothing

            If Not perRiferimento.TryGetValue(riferimentoBase, documento) Then
                Dim dettaglio As String = ValoreTesto(rigaExcel, 3).Trim()         'D

                documento = New DocumentoImpegnoImport() With {
                    .Riferimento = riferimentoBase &
                                   If(dettaglio.Length > 0, " " & dettaglio, ""),
                    .ConsegnaPrecedente = dettaglio.Length > 0,
                    .Cliente = ValoreTesto(rigaExcel, 0).Trim(),                   'A
                    .CodDestEsterno = NTSCInt(ValoreTesto(rigaExcel, 5)),         'F
                    .DataBaseConsegna = LeggiDataBaseConsegna(
                                            rigaExcel(7), riferimentoBase)        'H
                }

                perRiferimento.Add(riferimentoBase, documento)
                documenti.Add(documento)
            End If

            documento.Righe.Add(New RigaImpegnoImport() With {
                .TipoRiga = ValoreTesto(rigaExcel, 8).Trim(),                     'I
                .CodArt = CreaCodArt(rigaExcel),                                  'J-K
                .Descrizione = ValoreTesto(rigaExcel, 11).Trim(),                 'L
                .Quantita = NTSCDec(ValoreTesto(rigaExcel, 12))                   'M
            })
        Next

        Return documenti
    End Function

    Private Function LeggiExcel(ByVal percorsoFile As String) As DataTable
        Dim risultato As New DataTable()

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

                For c As Integer = 0 To lettore.FieldCount - 1
                    risultato.Columns.Add("F" & c.ToString(), GetType(Object))
                Next

                While lettore.Read()
                    Dim riga As DataRow = risultato.NewRow()

                    For c As Integer = 0 To lettore.FieldCount - 1
                        Dim valore As Object = lettore.GetValue(c)
                        riga(c) = If(valore Is Nothing, DBNull.Value, valore)
                    Next

                    risultato.Rows.Add(riga)
                End While
            End Using
        End Using

        If risultato.Columns.Count < 15 Then
            Throw New Exception("Il file Excel non contiene tutte le colonne richieste (A-O).")
        End If

        Return risultato
    End Function

    Private Function CreaCodArt(ByVal riga As DataRow) As String
        Dim prefisso As String = ValoreTesto(riga, 9).Trim()     'J
        Dim codice As String = ValoreTesto(riga, 10).Trim()      'K

        If String.IsNullOrWhiteSpace(prefisso) Then Return codice
        If String.IsNullOrWhiteSpace(codice) Then Return prefisso

        Return prefisso.TrimEnd("-"c) & "-" & codice.TrimStart("-"c)
    End Function

    Private Function ValoreTesto(ByVal riga As DataRow, ByVal indice As Integer) As String
        If riga Is Nothing OrElse indice < 0 OrElse indice >= riga.ItemArray.Length Then Return ""
        If riga(indice) Is DBNull.Value Then Return ""
        Return NTSCStr(riga(indice))
    End Function

    Private Sub RisolviDestinazione(ByVal conto As Integer,
                                      ByVal codDestEsterno As Integer,
                                      ByVal dataBase As Date,
                                      ByVal consegnaPrecedente As Boolean,
                                      ByRef codDest As Integer,
                                      ByRef dataConsegna As Nullable(Of Date),
                                      ByVal noteNonBloccanti As List(Of String))
        Try
            codDest = 0

            Dim risultato As DataTable =
                oClfGsor.GetDestinazioneImport(codDestEsterno.ToString(), conto)

            If risultato Is Nothing OrElse risultato.Rows.Count = 0 Then
                noteNonBloccanti.Add(
                    "Destinazione esterna " & codDestEsterno.ToString() &
                    " non trovata: utilizzata la destinazione principale (0).")
                Return
            End If

            codDest = NTSCInt(risultato.Rows(0)!dd_coddest)

            Dim giornoConsegna As String =
                NTSCStr(risultato.Rows(0)!dd_hhGiornoConsegna).Trim()

            If String.IsNullOrWhiteSpace(giornoConsegna) Then Return

            dataConsegna =
                CalcolaDataConsegna(giornoConsegna, dataBase, consegnaPrecedente)

            If Not dataConsegna.HasValue Then
                noteNonBloccanti.Add(
                    "Giorno consegna non riconosciuto per destinazione " &
                    codDestEsterno.ToString() & ": " & giornoConsegna)
            End If

        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
            codDest = 0
            noteNonBloccanti.Add(
                "Errore nella ricerca della destinazione esterna " &
                codDestEsterno.ToString() &
                ": utilizzata la destinazione principale (0).")
        End Try
    End Sub

    Private Function CalcolaDataConsegna(ByVal giorno As String,
                                         ByVal dataBase As Date,
                                         ByVal consegnaPrecedente As Boolean) As Nullable(Of Date)
        Dim valore As String =
            NTSCStr(giorno).Trim().ToLowerInvariant().Replace("ì", "i")

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

        'Colonna D valorizzata:
        'cerca il giorno di consegna precedente.
        If consegnaPrecedente Then

            Dim giorniDaSottrarre As Integer =
        (CInt(dataBase.DayOfWeek) - CInt(giornoSettimana) + 7) Mod 7

            'Se il giorno coincide, "precedente" significa
            'lo stesso giorno della settimana precedente.
            If giorniDaSottrarre = 0 Then
                giorniDaSottrarre = 7
            End If

            Return dataBase.Date.AddDays(-giorniDaSottrarre)

        End If

        'Regola già presente nella precedente implementazione:
        'se la data base è lunedì e la consegna è venerdì, usa il venerdì precedente.
        If dataBase.DayOfWeek = DayOfWeek.Monday AndAlso
           giornoSettimana = DayOfWeek.Friday Then
            Return dataBase.Date.AddDays(-3)
        End If

        Dim giorniDaAggiungere As Integer =
            (CInt(giornoSettimana) - CInt(dataBase.DayOfWeek) + 7) Mod 7

        Return dataBase.Date.AddDays(giorniDaAggiungere)
    End Function

    Private Function LeggiDataBaseConsegna(ByVal valore As Object,
                                           ByVal riferimento As String) As Date
        If TypeOf valore Is DateTime Then Return CType(valore, DateTime).Date
        If TypeOf valore Is Double Then Return DateTime.FromOADate(CDbl(valore)).Date

        Dim data As DateTime

        If valore IsNot Nothing AndAlso
           valore IsNot DBNull.Value AndAlso
           DateTime.TryParse(
               NTSCStr(valore),
               CultureInfo.GetCultureInfo("it-IT"),
               DateTimeStyles.None,
               data) Then

            Return data.Date
        End If

        Throw New Exception(
            "Data non valida nella colonna H per l'impegno " & riferimento & ".")
    End Function

    Private Function ConfermaReimportazione(ByVal riferimento As String,
                                            ByVal esistenti As DataTable) As Boolean
        Dim documenti As New List(Of String)()

        For Each riga As DataRow In esistenti.Rows
            Dim serie As String = NTSCStr(riga!td_serie).Trim()
            If serie = "" Then serie = "(vuota)"

            documenti.Add(
                "- Tipo " & NTSCStr(riga!td_tipork).Trim() &
                ", anno " & NTSCStr(riga!td_anno) &
                ", serie " & serie &
                ", numero " & NTSCStr(riga!td_numord) &
                ", conto " & NTSCStr(riga!td_conto))
        Next

        Dim messaggio As String =
            "Il riferimento " & riferimento &
            " è già presente nei seguenti documenti:" & vbCrLf &
            String.Join(vbCrLf, documenti.ToArray()) & vbCrLf & vbCrLf &
            "Sei sicuro di voler reimportare questo ordine?"

        Return oApp.MsgBoxInfoYesNo_DefNo(messaggio) = DialogResult.Yes
    End Function

    Private Function PercentualeDocumento(ByVal indiceDocumento As Integer,
                                          ByVal totaleDocumenti As Integer,
                                          ByVal rigaElaborata As Integer,
                                          ByVal totaleRighe As Integer) As Integer
        If totaleDocumenti <= 0 Then Return 5

        Dim quotaDocumento As Decimal = 90D / totaleDocumenti
        Dim avanzamentoRighe As Decimal = 0D

        If totaleRighe > 0 Then
            avanzamentoRighe = quotaDocumento * rigaElaborata / totaleRighe
        End If

        Return CInt(Math.Truncate(
            5D + quotaDocumento * (indiceDocumento - 1) + avanzamentoRighe))
    End Function

    Private Function SelezionaFileImpegni() As String
        Try
            Using dialogo As New OpenFileDialog()
                dialogo.Title = "Seleziona il file degli impegni clienti"
                dialogo.Filter = "File Excel 97-2003 (*.xls)|*.xls|Tutti i file (*.*)|*.*"
                dialogo.CheckFileExists = True
                dialogo.Multiselect = False

                If Not String.IsNullOrWhiteSpace(strUltimaCartellaImport) AndAlso
                   Directory.Exists(strUltimaCartellaImport) Then
                    dialogo.InitialDirectory = strUltimaCartellaImport
                End If

                If dialogo.ShowDialog(Me) <> DialogResult.OK Then Return ""

                strUltimaCartellaImport = Path.GetDirectoryName(dialogo.FileName)
                Return dialogo.FileName
            End Using

        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
            Return ""
        End Try
    End Function

    Private Sub MostraEsitoImportazione(ByVal esito As EsitoImportazioneImpegni)
        Try
            If esito Is Nothing Then
                oApp.MsgBoxErr("Importazione impegni non completata.")
                Return
            End If

            Dim riepilogo As String =
                "Impegni salvati: " & esito.DocumentiSalvati.ToString() & "."

            If esito.DocumentiSaltati > 0 Then
                riepilogo &= " Reimportazioni annullate: " &
                             esito.DocumentiSaltati.ToString() & "."
            End If

            If esito.ErroriDocumento.Count = 0 Then
                oApp.MsgBoxInfo("Importazione completata. " & riepilogo)
                Return
            End If

            Dim messaggio As String =
                riepilogo & vbCrLf &
                "Documenti non salvati:" & vbCrLf &
                String.Join(vbCrLf, esito.ErroriDocumento.ToArray())

            oApp.MsgBoxExclamation(messaggio)

        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
        End Try
    End Sub

    Private Function GetGlyphImportaImpegni() As Image
        Try
            Using stream As Stream =
                GetType(FROORGSOR).Assembly.GetManifestResourceStream(
                    "NTSInformatica.ImportaImpegni24.png")

                If stream Is Nothing Then Return Nothing

                Using img As Image = Image.FromStream(stream)
                    Return New Bitmap(img)
                End Using
            End Using

        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
            Return Nothing
        End Try
    End Function

#End Region

End Class
