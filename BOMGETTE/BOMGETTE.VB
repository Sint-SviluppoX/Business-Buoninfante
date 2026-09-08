Imports System.Data
Imports NTSInformatica.CLN__STD
Imports System
Public Class FROMGETTE
  Inherits FRMMGETTE
  Public _oClfEtte As CLFMGETTE
  Public WithEvents cb_hhstamparifd As NTSInformatica.NTSComboBox
  Public Property oClfEtte() As CLFMGETTE
    Get
      If _oClfEtte Is Nothing Then _oClfEtte = CType(oCleEtte, CLFMGETTE)
      Return _oClfEtte
    End Get
    Set(ByVal value As CLFMGETTE)
      _oClfEtte = value
    End Set
  End Property
  Public Overrides Sub InitControls()
    MyBase.InitControls()
    Try
      cb_hhstamparifd = CType(NTSFindControlByName(Me, "cb_hhstamparfid"), NTSComboBox)
    Catch ex As Exception
      CLN__STD.GestErr(ex, Me, "")
    End Try
  End Sub
  Public Overrides Sub Stampa(ByVal nDestin As Integer)
    Dim nPjob As Object
    Dim nRis As Integer = 0
    Dim strCrpe As String = ""
    Dim i As Integer
    Dim j As Integer
    Dim x As Integer
    Dim strNomRpt As String = ""
    Dim strKey As String = ""
    Dim lIIstMatrTmp As Integer
    Dim bDocMagazz As Boolean = False
    Dim bDaDoc As Boolean = False
    Dim strCode As String = ""
    Dim strCodartList As String = ""
    Dim strCodart As String = ""
    Dim bAvvisare As Boolean = True
    Dim dttOut As New DataTable
    Dim dttTmp As New DataTable
    Dim dttTmp2 As New DataTable
    Dim frmEtba As FRMMGETBA = Nothing
    Dim nNEtichettePerTC As Integer = 0
    Dim oDttgr As New CLEGROUPBY
    Dim dttGr As New DataTable
    Dim bOperazioneAnnullata As Boolean = False
    Dim strTipodoc As String = "".PadLeft(1)
    Dim bContinua As Boolean = False

    Try
      ' Valido l'ultimo dato inserito
      Me.ValidaLastControl()

      If Not ControlliAggiuntivi() Then Return

      oMenu.ResetTblInstId("TTSTMATS", bsDbAzienda, lIIstMats)
      oMenu.ResetTblInstId("TTSTMATR", bsDbAzienda, lIIstMatr)

      lIIstMatrTmp = oMenu.GetTblInstId("TTSTMATR", bsDbAzienda)
      '--- Controlla barcode mancanti

      '-----------------------------------------------------------------------------------------
      '--- Riempie la tabella temporanea
      System.Windows.Forms.Cursor.Current = Cursors.WaitCursor

      If Not oCleEtte.LeggiTable(cbOrigine.SelectedValue, bBD_DaCodarfo, edArtico.Text, strArticoli, strWhereArtico, NTSCInt(edLista.Text), cbTipodoc.SelectedValue,
                                 bCreaBarcode, NTSCInt(edDaclie.Text), NTSCInt(edAclie.Text), NTSCDate(edDadata.Text), NTSCDate(edAdata.Text),
                                 NTSCInt(edDaanno.Text), NTSCInt(edAanno.Text), edDaserie.Text, edAserie.Text, NTSCInt(edDanumero.Text),
                                 NTSCInt(edAnumero.Text), NTSCInt(edDariga.Text), NTSCInt(edAriga.Text), opColli.Checked, dttOut, dttTmp, dttTmp2,
                                 bDocMagazz, bDaDoc, strCode, strCodartList, strCodart, bAvvisare, ckGenera.Checked, cbTipobar.SelectedValue, cbTipoDoc2.SelectedValue,
                                 ckDatamod.Checked, NTSCDate(edDatamod.Text)) Then Return
      '---------------------------------------------------------------------------------------------------------
      If oCallParams IsNot Nothing AndAlso
         oCallParams.strNomProg.ToUpper = "BNDKKONS" Then GoTo SaltaRichiestaDiStampa
      '---------------------------------------------------------------------------------------------------------
      If dttOut IsNot Nothing AndAlso
         dttOut.Rows.Count > 0 Then
        If (NTSCStr(cbOrigine.SelectedValue) = "A" OrElse
            NTSCStr(cbOrigine.SelectedValue) = "L") AndAlso
           oApp.MsgBoxInfoYesNo_DefNo(oApp.Tr(Me, 132748885484015332, "Attenzione!" & vbCrLf &
             "Saranno trattati/stampati |" & Format(NTSCDec(dttOut.Rows.Count), "#,##0") & "| articoli." & vbCrLf &
             "Procedere?")) = DialogResult.No Then
          bOperazioneAnnullata = True
        ElseIf NTSCStr(cbOrigine.SelectedValue) = "M" AndAlso
               oDttgr.NTSGroupBy(dttOut, dttGr, "tipork, anno, serie, numero", "", "tipork, anno, serie, numero") = True AndAlso
               dttGr.Rows.Count > 0 AndAlso
               oApp.MsgBoxInfoYesNo_DefNo(oApp.Tr(Me, 132748888977380574, "Attenzione!" & vbCrLf &
               "Saranno trattati/stampati |" & Format(NTSCDec(dttGr.Rows.Count), "#,##0") & "| documenti." & vbCrLf &
               "Procedere?")) = DialogResult.No Then
          bOperazioneAnnullata = True
        End If
        If bOperazioneAnnullata = True Then
          oApp.MsgBoxInfo(oApp.Tr(Me, 132748886144908324, "Operazione annullata."))
          Return
        End If
      End If
      '---------------------------------------------------------------------------------------------------------
SaltaRichiestaDiStampa:
      '---------------------------------------------------------------------------------------------------------
      '--- Se:
      '--- Origine dati = "Lista selezionata articoli"
      '--- E' indicata una taglia
      '--- Il CheckBox "Tutte le taglie gestite dall'articolo" NON è selezionato
      '--- elimina, dal DataTable di ritorno "dttOut" tutti gli articoli che non appartengono alla taglia indicata.
      '--- Se, alla fine non esistono dati, avvisa ed esce
      '---------------------------------------------------------------------------------------------------------
      If dttOut.Rows.Count > 0 AndAlso
         NTSCStr(cbOrigine.SelectedValue) = "L" AndAlso
         edTaglia.Text.Trim <> "" AndAlso
         ckTutteLeTaglie.Checked = False Then
        Dim dtrT() As DataRow = dttOut.Select("lsa_tctaglia <> " & CStrSQL(edTaglia.Text))
        If dtrT.Length > 0 Then
          For i = 0 To (dtrT.Length - 1)
            dtrT(i).Delete()
          Next
        End If
        dttOut.AcceptChanges()
        If dttOut.Rows.Count = 0 Then
          oApp.MsgBoxInfo(oApp.Tr(Me, 131890981775378577, "Attenzione!" & vbCrLf &
            "Non esistono articoli da stampare con le caratteristiche richieste."))
          Return
        End If
      End If
      '-----------------------------------------------------------------------------------------
      '--- Esecuzione della finestra sui Barcode
      '-----------------------------------------------------------------------------------------
      'Controllo se esistono dei codici a barre
      If bNEtichetteBarCode Then
        For i = 0 To dttOut.Rows.Count - 1
          dttTmp = Nothing
          If Not oCleEtte.SelectBarcode(ckBarcode.Checked, cbSolobar.SelectedValue, NTSCStr(dttOut.Rows(i)!Codart), dttTmp) Then Return

          'se è stata selezionata la ckDatamod, non stampo le etichette con data ultima modifica >= della data indicata in edDatamod (WSS 326842)
          If ckDatamod.Checked Then
            For l As Integer = (dttTmp.Rows.Count - 1) To 0 Step -1
              If NTSCDate(dttTmp.Rows(l)!bc_datagg) < NTSCDate(edDatamod.Text) Then dttTmp.Rows(l).Delete()
            Next
            dttTmp.AcceptChanges()
          End If

          If dttTmp.Rows.Count <> 0 Then
            bContinua = True
            If Not oCleEtte.SelectTTstmatr(lIIstMatrTmp, NTSCStr(dttOut.Rows(i)!Codart), dttTmp2) Then Return

            If dttTmp2.Rows.Count = 0 Then
              If dttTmp.Rows.Count > 1 Then
                For x = 0 To dttTmp.Rows.Count - 1
                  If Not oCleEtte.InsertBarcodeTTstmatr(lIIstMatrTmp, NTSCStr(dttOut.Rows(i)!Codart), NTSCStr(dttTmp.Rows(x)!bc_code),
                                                        NTSCDec(dttTmp.Rows(x)!bc_quant), NTSCStr(dttTmp.Rows(x)!bc_unmis)) Then Return
                Next

                frmEtba = CType(NTSNewFormModal("FRMMGETBA"), FRMMGETBA)
                If Not frmEtba.Init(oMenu, Nothing, DittaCorrente, Nothing) Then Return
                frmEtba.lIIstMatrTmp = lIIstMatrTmp
                frmEtba.InitEntity(oCleEtte)
                frmEtba.strCodart = NTSCStr(dttOut.Rows(i)!Codart)
                '----------------------------------------------------
                'WSS 327276: la FRMMGETBA mi farà vedere in 'lbArticolo' di quale articolo sono i codici a barre che voglio stampare
                frmEtba.lbArticolo.Text = "Articolo: " & NTSCStr(dttOut.Rows(i)!Codart) & " - " & NTSCStr(dttOut.Rows(i)!Descr)
                '----------------------------------------------------
                frmEtba.ShowDialog(Me)

                If Not frmEtba.bOk Then Return
              End If
            End If
          End If
        Next
      Else
        bContinua = True
      End If

      'se per tutti gli articoli non sono presenti etichette con data ultima modifica >= della data indicata in edDatamod, non eseguo la stampa (WSS 326842)
      If ckDatamod.Checked AndAlso bContinua = False Then
        oApp.MsgBoxErr(oApp.Tr(Me, 132990734447396115, "Non sono presenti articoli con prezzo di listino modificato dal " & edDatamod.Text & "."))
        Return
      End If

      Me.Refresh()


      'Tipo: E=ean13, C=Code39, 2=2su5 interleaved, 8=Ean8,I=ISBN, F=farmaceutico, A=AltroIndustr., U=UPC
      'Pers. Alfy 06-04-21
      '---------------------------------------------------------
      'Solo per Ordini di produzione
      'If cbOrigine.SelectedValue = "M" AndAlso cbTipodoc.SelectedValue = "H" AndAlso cb_hhstamparifd.Enabled Then
      If cbOrigine.SelectedValue = "M" AndAlso cbTipodoc.SelectedValue = "H" Then
        'Se ho scelto la stampa diversa da "solo carta", procedo!
        If cb_hhstamparifd.SelectedValue <> "RPT" Then
          Dim dttTmpRfid As New DataTable
          'Carico un mio datatable con solo le righe presi dai documenti filtrati e solo quelle che hanno l'apposito campo personalizzato a "S"
          If oClfEtte.SelezionaDaDocMagaRFID(oApp.Ditta, cbTipodoc.SelectedValue.ToString, NTSCInt(edDaclie.Text), NTSCInt(edAclie.Text), NTSCDate(edDadata.Text), NTSCDate(edAdata.Text), NTSCInt(edDaanno.Text), NTSCInt(edAanno.Text), edDaserie.Text, edAserie.Text, NTSCInt(edDanumero.Text),
                                 NTSCInt(edAnumero.Text), NTSCInt(edDariga.Text), NTSCInt(edAriga.Text), bDocMagazz, opColli.Checked, dttTmpRfid) Then
            Dim NRecordDelete As Integer = 0
            Dim TmpDesArt As String = String.Empty
            Dim TmpQtaEtic As Integer = 0
            Dim RisQuestDelete As DialogResult = DialogResult.Cancel
            'If oApp.MsgBoxInfoRedYesNo_DefNo("Cancellare tutte le etichette RFID caricate prima di oggi (" & Now.Date.ToString("dd/MM/yyyy") & ")?") = DialogResult.Yes Then oClfEtte.CancellaEtichettePrecedentiADataDBExt(Now.Date, "S", NRecordDelete)
            RisQuestDelete = oApp.MsgBoxInfoYesNoCancel_DefCancel("Svuotare tutte le etichette RFID caricate in precedenza?" & Chr(13) & "SI = (Tutte)" & Chr(13) & "NO = (Solo quelle caricate prima del " & Now.Date.ToString("dd/MM/yyyy") & ")" & Chr(13) & "ANNULLA = (No)")
            'Svuoto tutta la tabella
            If RisQuestDelete = DialogResult.Yes Then oClfEtte.CancellaEtichetteAllDBExt(NRecordDelete)
            'Cancello solo quelle precedenti alla data
            If RisQuestDelete = DialogResult.No Then oClfEtte.CancellaEtichettePrecedentiADataDBExt(Now.Date, "S", NRecordDelete)
            'Controllo che ci sono righe
            If dttTmpRfid.Rows.Count > 0 Then
              TmpDesArt = String.Empty
              TmpQtaEtic = 0
              'Ciclo le righe e cancello e inserisco riga in DB esterno con info articolo, info documento e quantità
              For iEticDB As Integer = 0 To dttTmpRfid.Rows.Count - 1
                'Lancio una delete per cancellare eventuale riga creata in precedenza con le stesse chiavi
                oClfEtte.CancellaEtichetteDBExt(dttTmpRfid.Rows(iEticDB)("tipork").ToString, NTSCInt(dttTmpRfid.Rows(iEticDB)("anno").ToString), dttTmpRfid.Rows(iEticDB)("serie").ToString, NTSCInt(dttTmpRfid.Rows(iEticDB)("numero").ToString), NTSCInt(dttTmpRfid.Rows(iEticDB)("riga").ToString))
                TmpDesArt = String.Empty
                TmpDesArt = oClfEtte.MyValCodiceDB(dttTmpRfid.Rows(iEticDB)("codart").ToString, "ARTICO", "S")

                dttTmp = Nothing
                If Not oCleEtte.SelectBarcode(ckBarcode.Checked, cbSolobar.SelectedValue, NTSCStr(dttTmpRfid.Rows(iEticDB)("codart").ToString), dttTmp) Then Return
                'Se l'articolo esiste e anche l'ean13...allora procedo.
                If Not String.IsNullOrEmpty(TmpDesArt) AndAlso dttTmp.Rows.Count > 0 Then
                  TmpQtaEtic += CInt(dttTmpRfid.Rows(iEticDB)("quant").ToString)
                  oClfEtte.InserisciEtichetteDBExt(dttTmpRfid.Rows(iEticDB)("tipork").ToString, NTSCInt(dttTmpRfid.Rows(iEticDB)("anno").ToString), dttTmpRfid.Rows(iEticDB)("serie").ToString, NTSCInt(dttTmpRfid.Rows(iEticDB)("numero").ToString), NTSCInt(dttTmpRfid.Rows(iEticDB)("riga").ToString), NTSCDate(dttTmpRfid.Rows(iEticDB)("dtdoc").ToString), dttTmp.Rows(dttTmp.Rows.Count - 1)("bc_code").ToString, TmpDesArt, CDbl(dttTmpRfid.Rows(iEticDB)("quant").ToString), TmpQtaEtic)
                  'oClfEtte.InserisciEtichetteDBExt(dttTmpRfid.Rows(iEticDB)("tipork").ToString, NTSCInt(dttTmpRfid.Rows(iEticDB)("anno").ToString), dttTmpRfid.Rows(iEticDB)("serie").ToString, NTSCInt(dttTmpRfid.Rows(iEticDB)("numero").ToString), NTSCInt(dttTmpRfid.Rows(iEticDB)("riga").ToString), NTSCDate(dttTmpRfid.Rows(iEticDB)("dtdoc").ToString), dttTmpRfid.Rows(iEticDB)("codart").ToString, TmpDesArt, CDbl(dttTmpRfid.Rows(iEticDB)("quant").ToString), TmpQtaEtic)
                End If
              Next
              oApp.MsgBoxInfo("Operazione completata!" & Chr(13) & "Inviate N° " & TmpQtaEtic & " etichette RFID!" & Chr(13) & IIf(NRecordDelete > 0, "Cancellate N° " & NRecordDelete & " righe di etichette RFID!", String.Empty).ToString, True)
            End If
            'Se non ho messo "solo RFID" continuo con il resto (standard)
            If cb_hhstamparifd.SelectedValue = "RFI" Then
              oApp.MsgBoxInfo("Operazione completata!" & Chr(13) & "Inviate N° " & TmpQtaEtic & " etichette RFID!" & Chr(13) & IIf(NRecordDelete > 0, "Cancellate N° " & NRecordDelete & " righe di etichette RFID!", String.Empty).ToString)
              Return
            End If
          End If
        End If
      End If
      '---------------------------------------------------------


      If edNEtichettePerTC.Visible = True Then nNEtichettePerTC = NTSCInt(edNEtichettePerTC.Text)

      If Not oCleEtte.RiempiTmpTable(ckPrezzi.Checked, ckDatamod.Checked, NTSCDate(edDatamod.Text), ckBarcode.Checked, cbSolobar.SelectedValue,
                                     cbBarcode.SelectedValue, ckArtbar.Checked, cbOrigine.SelectedValue, bBD_DaCodarfo, NTSCInt(edLista.Text), opColli.Checked,
                                     bNEtichetteBarCode, ckValuta.Checked, NTSCInt(edListino.Text), NTSCDate(edValida.Text), NTSCInt(edValuta.Text),
                                     ckStampe.Checked, cbOperazione.SelectedValue, ckMisura.Checked, NTSCInt(edNumMolt.Text), bDividiQuantConfez,
                                     lIIstMatr, lIIstMats, bUsaTtstmatr, NTSCInt(edEtiche.Text), nBD_Listino, bModTCO,
                                     IIf(ckTutteLeTaglie.Checked = False, edTaglia.Text, "".PadLeft(5, "*"c)).ToString,
                                     lIIstMatrTmp, bStampaNegativi, dttOut, dttTmp, dttTmp2, bDocMagazz,
                                     bDaDoc, strCode, strCodartList, strCodart, bAvvisare, nNEtichettePerTC, ckNumeti.Checked, ckPrzNetClasseSconto.Checked,
                                     opNumetiDaNote.Checked) Then Return
      '------------------------------------------------------------------------------------------------------------
      '--- Se chiamato da voce di menu/strumenti "Stampa etichette finale" in 
      '--- Gestione Documenti di Magazzino (BNVEBOLL) o Gestione Ordini/Impegni/Preventivi (BNORGSOR)
      '--- richiama una griglia, raggruppata per articolo, dove si possono modificare le quantità da stampare
      '------------------------------------------------------------------------------------------------------------
      If ModificaNumeroEtichetteDaStampare() = False Then Return
      '------------------------------------------------------------------------------------------------------------
      If ckTermica.Checked And nDestin = 1 Then
        If bDometic Then
          StampaDirettaDometic()
        Else
          '--- Stampo direttamente sulla stampante termica senza usare Crystal report
          StampaDiretta()
        End If
        Return
      End If
      '-----------------------------------------------------------------------------------------
      '--- determina il nome del report da stampare e la cartella relativa del registro
      '-----------------------------------------------------------------------------------------
      SettaNomeReport(strKey, strNomRpt)
      '------------------------------------------------------------------------------------------------------------
      '--- Vedi valutazione parametri d'ingresso nella "FRMMGETTE_ActivatedFirst"
      '------------------------------------------------------------------------------------------------------------
      If oCallParams IsNot Nothing AndAlso
         (oCallParams.strNomProg.ToUpper = "BNORGSOR" OrElse
          oCallParams.strNomProg.ToUpper = "BNVEBOLL" OrElse
          oCallParams.strNomProg.ToUpper = "BNDKKONS") AndAlso
         oCallParams.bFlag1 = True AndAlso
         oCallParams.bPar1 = True Then oMenu.ReportSetprinter = False
      '-----------------------------------------------------------------------------------------
      '--- Lancia le stampe
      '-----------------------------------------------------------------------------------------
      If NTSCStr(cbOrigine.SelectedValue) = "H" OrElse
         NTSCStr(cbOrigine.SelectedValue) = "M" Then strTipodoc = NTSCStr(cbTipodoc.SelectedValue)
      nPjob = oMenu.ReportPEInit(oApp.Ditta, Me, "BSMGETTE", strKey, strTipodoc, 0, nDestin, strNomRpt, False, "Stampa etichette", False)
      '-----------------------------------------------------------------------------------------
      '--- Se la stampa è stata annullata esce
      '-----------------------------------------------------------------------------------------
      If nPjob Is Nothing Then Return
      '-----------------------------------------------------------------------------------------
      '--- Formula di selezione
      '-----------------------------------------------------------------------------------------
      If bUsaTtstmatr And (cbOrigine.SelectedValue = "M" Or cbOrigine.SelectedValue = "H") Then
        strCrpe = "{artico.codditt} = '" & DittaCorrente & "'" &
                    " And {ttstmatr.codditt} = '" & DittaCorrente & "'" &
                    " And {ttstmatr.instid} = " & lIIstMatr
      Else
        strCrpe = "{artico.codditt} = '" & DittaCorrente & "'" &
                    " And {ttstmats.codditt} = '" & DittaCorrente & "'" &
                    " And {ttstmats.instid} = " & lIIstMats
      End If
      '------------------------------------------------------------------------------------------------------------
      'lancio tutti gli eventuali reports (gestisce già il multireport)
      For i = 1 To UBound(CType(nPjob, Array), 2)
        If i = 1 Then
          nRis = oMenu.PESetFormula(NTSCInt(CType(nPjob, Array).GetValue(0, i)), "STAMPAPREZZI", "'" & NTSCStr(IIf(ckPrezzi.Checked, "S", "N")) & "'")
          If ckPrezzi.Checked Then
            nRis = oMenu.PESetFormula(NTSCInt(CType(nPjob, Array).GetValue(0, i)), "DATAVAL", "'" & edValida.Text & "'")
            nRis = oMenu.PESetFormula(NTSCInt(CType(nPjob, Array).GetValue(0, i)), "LISTINO", edListino.Text)
            nRis = oMenu.PESetFormula(NTSCInt(CType(nPjob, Array).GetValue(0, i)), "STAMPAPREZZIVAL", "'" & NTSCStr(IIf(ckValuta.Checked, "S", "N")) & "'")
            If ckValuta.Checked Then
              nRis = oMenu.PESetFormula(NTSCInt(CType(nPjob, Array).GetValue(0, i)), "CODVALU", edValuta.Text)
              nRis = oMenu.PESetFormula(NTSCInt(CType(nPjob, Array).GetValue(0, i)), "DESVALU", "'" & lbDescvalu.Text & "'")
            End If
          End If
          nRis = oMenu.PESetSelectionFormula(NTSCInt(CType(nPjob, Array).GetValue(0, i)), strCrpe)
          nRis = oMenu.ReportPEVai(NTSCInt(CType(nPjob, Array).GetValue(0, i)))
        Else
          If Trim$(NTSCStr(CType(nPjob, Array).GetValue(2, i))) <> "" Then
            'RecordSelectionFormula
            nRis = oMenu.PESetSelectionFormula(NTSCInt(CType(nPjob, Array).GetValue(0, i)), CrpeResolveFormula(Me, NTSCStr(CType(nPjob, Array).GetValue(2, i)), strCrpe))
          End If
          nRis = oMenu.PESetFormula(NTSCInt(CType(nPjob, Array).GetValue(0, i)), "STAMPAPREZZI", "'" & NTSCStr(IIf(ckPrezzi.Checked, "S", "N")) & "'")
          If ckPrezzi.Checked Then
            nRis = oMenu.PESetFormula(NTSCInt(CType(nPjob, Array).GetValue(0, i)), "DATAVAL", edValida.Text)
            nRis = oMenu.PESetFormula(NTSCInt(CType(nPjob, Array).GetValue(0, i)), "LISTINO", edListino.Text)
            nRis = oMenu.PESetFormula(NTSCInt(CType(nPjob, Array).GetValue(0, i)), "STAMPAPREZZIVAL", "'" & NTSCStr(IIf(ckValuta.Checked, "S", "N")) & "'")
            If ckValuta.Checked Then
              nRis = oMenu.PESetFormula(NTSCInt(CType(nPjob, Array).GetValue(0, i)), "CODVALU", edValuta.Text)
              nRis = oMenu.PESetFormula(NTSCInt(CType(nPjob, Array).GetValue(0, i)), "DESVALU", "'" & lbDescvalu.Text & "'")
            End If
          End If
          For j = 3 To 12
            'FormulaJ
            If Trim$(NTSCStr(CType(nPjob, Array).GetValue(j, i))) <> "" Then
              nRis = oMenu.PESetFormula(CInt(CType(nPjob, Array).GetValue(0, i)), NTSCStr(CType(nPjob, Array).GetValue(j, i)), CrpeResolveFormula(Me, NTSCStr(CType(nPjob, Array).GetValue(j + 10, i))))
            End If
          Next j
          nRis = oMenu.ReportPEVai(NTSCInt(CType(nPjob, Array).GetValue(0, i)))
        End If
      Next

    Catch ex As Exception
      '-------------------------------------------------
      CLN__STD.GestErr(ex, Me, "")
      '-------------------------------------------------
    Finally
      '------------------------------------------------------------------------------------------------------------
      oDttgr = Nothing
      dttGr.Clear() : dttGr.Dispose()
      '------------------------------------------------------------------------------------------------------------
      If Not frmEtba Is Nothing Then frmEtba.Dispose()
      frmEtba = Nothing
      System.Windows.Forms.Cursor.Current = Cursors.Default
      lbAvviso.Text = oApp.Tr(Me, 130421093076039252, "Nessuna operazione in corso.")
      '------------------------------------------------------------------------------------------------------------
      oMenu.ReportSetprinter = True
      '------------------------------------------------------------------------------------------------------------
    End Try
  End Sub
  Public Overrides Sub cbTipodoc_EditValueChanged(sender As Object, e As EventArgs)
    Try
      MyBase.cbTipodoc_EditValueChanged(sender, e)
      'Pers.
      '--------------------------------------------
      cb_hhstamparifd.Enabled = False
      If cbTipodoc.SelectedValue = "H" Then
        cb_hhstamparifd.Enabled = True
        cb_hhstamparifd.SelectedValue = "RFI"
        cb_hhstamparifd.Enabled = False
        'Metto fissi questi campi (04-06-21)
        ckArtbar.Checked = True
        ckBarcode.Checked = True
        cbSolobar.SelectedValue = "E"
        ckArtbar.Enabled = False
        ckBarcode.Enabled = False
        cbSolobar.Enabled = False
      End If
      '--------------------------------------------
    Catch ex As Exception
      CLN__STD.GestErr(ex, Me, "")
    End Try
  End Sub

#Region "PreSR8Cu4"
  'Public Overrides Sub Stampa(ByVal nDestin As Integer)
  '  Dim nPjob As Object
  '  Dim nRis As Integer = 0
  '  Dim strCrpe As String = ""
  '  Dim i As Integer
  '  Dim j As Integer
  '  Dim x As Integer
  '  Dim strNomRpt As String = ""
  '  Dim strKey As String = ""
  '  Dim lIIstMatrTmp As Integer
  '  Dim bDocMagazz As Boolean = False
  '  Dim bDaDoc As Boolean = False
  '  Dim strCode As String = ""
  '  Dim strCodartList As String = ""
  '  Dim strCodart As String = ""
  '  Dim bAvvisare As Boolean = True
  '  Dim dttOut As New DataTable
  '  Dim dttTmp As New DataTable
  '  Dim dttTmp2 As New DataTable
  '  Dim frmEtba As FRMMGETBA = Nothing
  '  Dim nNEtichettePerTC As Integer = 0
  '  Dim oDttgr As New CLEGROUPBY
  '  Dim dttGr As New DataTable
  '  Dim bOperazioneAnnullata As Boolean = False

  '  Try
  '    ' Valido l'ultimo dato inserito
  '    Me.ValidaLastControl()

  '    If Not ControlliAggiuntivi() Then Return

  '    oMenu.ResetTblInstId("TTSTMATS", bsDbAzienda, lIIstMats)
  '    oMenu.ResetTblInstId("TTSTMATR", bsDbAzienda, lIIstMatr)

  '    lIIstMatrTmp = oMenu.GetTblInstId("TTSTMATR", bsDbAzienda)
  '    '--- Controlla barcode mancanti

  '    '-----------------------------------------------------------------------------------------
  '    '--- Riempie la tabella temporanea
  '    System.Windows.Forms.Cursor.Current = Cursors.WaitCursor

  '    If Not oCleEtte.LeggiTable(cbOrigine.SelectedValue, bBD_DaCodarfo, edArtico.Text, strArticoli, strWhereArtico, NTSCInt(edLista.Text), cbTipodoc.SelectedValue,
  '                               bCreaBarcode, NTSCInt(edDaclie.Text), NTSCInt(edAclie.Text), NTSCDate(edDadata.Text), NTSCDate(edAdata.Text),
  '                               NTSCInt(edDaanno.Text), NTSCInt(edAanno.Text), edDaserie.Text, edAserie.Text, NTSCInt(edDanumero.Text),
  '                               NTSCInt(edAnumero.Text), NTSCInt(edDariga.Text), NTSCInt(edAriga.Text), opColli.Checked, dttOut, dttTmp, dttTmp2,
  '                               bDocMagazz, bDaDoc, strCode, strCodartList, strCodart, bAvvisare, ckGenera.Checked, cbTipobar.SelectedValue, cbTipoDoc2.SelectedValue) Then Return
  '    '---------------------------------------------------------------------------------------------------------
  '    If dttOut IsNot Nothing AndAlso
  '       dttOut.Rows.Count > 0 Then
  '      If (NTSCStr(cbOrigine.SelectedValue) = "A" OrElse
  '          NTSCStr(cbOrigine.SelectedValue) = "L") AndAlso
  '         oApp.MsgBoxInfoYesNo_DefNo(oApp.Tr(Me, 132748885484015332, "Attenzione!" & vbCrLf &
  '           "Saranno trattati/stampati |" & Format(NTSCDec(dttOut.Rows.Count), "#,##0") & "| articoli." & vbCrLf &
  '           "Procedere?")) = DialogResult.No Then
  '        bOperazioneAnnullata = True
  '      ElseIf NTSCStr(cbOrigine.SelectedValue) = "M" AndAlso
  '             oDttgr.NTSGroupBy(dttOut, dttGr, "tipork, anno, serie, numero", "", "tipork, anno, serie, numero") = True AndAlso
  '             dttGr.Rows.Count > 0 AndAlso
  '             oApp.MsgBoxInfoYesNo_DefNo(oApp.Tr(Me, 132748888977380574, "Attenzione!" & vbCrLf &
  '             "Saranno trattati/stampati |" & Format(NTSCDec(dttGr.Rows.Count), "#,##0") & "| documenti." & vbCrLf &
  '             "Procedere?")) = DialogResult.No Then
  '        bOperazioneAnnullata = True
  '      End If
  '      If bOperazioneAnnullata = True Then
  '        oApp.MsgBoxInfo(oApp.Tr(Me, 132748886144908324, "Operazione annullata."))
  '        Return
  '      End If
  '    End If
  '    '---------------------------------------------------------------------------------------------------------
  '    '--- Se:
  '    '--- Origine dati = "Lista selezionata articoli"
  '    '--- E' indicata una taglia
  '    '--- Il CheckBox "Tutte le taglie gestite dall'articolo" NON è selezionato
  '    '--- elimina, dal DataTable di ritorno "dttOut" tutti gli articoli che non appartengono alla taglia indicata.
  '    '--- Se, alla fine non esistono dati, avvisa ed esce
  '    '---------------------------------------------------------------------------------------------------------
  '    If dttOut.Rows.Count > 0 AndAlso
  '       NTSCStr(cbOrigine.SelectedValue) = "L" AndAlso
  '       edTaglia.Text.Trim <> "" AndAlso
  '       ckTutteLeTaglie.Checked = False Then
  '      Dim dtrT() As DataRow = dttOut.Select("lsa_tctaglia <> " & CStrSQL(edTaglia.Text))
  '      If dtrT.Length > 0 Then
  '        For i = 0 To (dtrT.Length - 1)
  '          dtrT(i).Delete()
  '        Next
  '      End If
  '      dttOut.AcceptChanges()
  '      If dttOut.Rows.Count = 0 Then
  '        oApp.MsgBoxInfo(oApp.Tr(Me, 131890981775378577, "Attenzione!" & vbCrLf &
  '          "Non esistono articoli da stampare con le caratteristiche richieste."))
  '        Return
  '      End If
  '    End If
  '    '-----------------------------------------------------------------------------------------
  '    '--- Esecuzione della finestra sui Barcode
  '    '-----------------------------------------------------------------------------------------
  '    'Controllo se esistono dei codici a barre
  '    If bNEtichetteBarCode Then
  '      For i = 0 To dttOut.Rows.Count - 1
  '        dttTmp = Nothing
  '        If Not oCleEtte.SelectBarcode(ckBarcode.Checked, cbSolobar.SelectedValue, NTSCStr(dttOut.Rows(i)!Codart), dttTmp) Then Return

  '        If dttTmp.Rows.Count <> 0 Then
  '          If Not oCleEtte.SelectTTstmatr(lIIstMatrTmp, NTSCStr(dttOut.Rows(i)!Codart), dttTmp2) Then Return

  '          If dttTmp2.Rows.Count = 0 Then
  '            If dttTmp.Rows.Count > 1 Then
  '              For x = 0 To dttTmp.Rows.Count - 1
  '                If Not oCleEtte.InsertBarcodeTTstmatr(lIIstMatrTmp, NTSCStr(dttOut.Rows(i)!Codart), NTSCStr(dttTmp.Rows(x)!bc_code),
  '                                                      NTSCDec(dttTmp.Rows(x)!bc_quant), NTSCStr(dttTmp.Rows(x)!bc_unmis)) Then Return
  '              Next

  '              frmEtba = CType(NTSNewFormModal("FRMMGETBA"), FRMMGETBA)
  '              If Not frmEtba.Init(oMenu, Nothing, DittaCorrente, Nothing) Then Return
  '              frmEtba.lIIstMatrTmp = lIIstMatrTmp
  '              frmEtba.InitEntity(oCleEtte)
  '              frmEtba.strCodart = NTSCStr(dttOut.Rows(i)!Codart)
  '              frmEtba.ShowDialog(Me)

  '              If Not frmEtba.bOk Then Return
  '            End If
  '          End If
  '        End If
  '      Next
  '    End If

  '    Me.Refresh()


  '    'Tipo: E=ean13, C=Code39, 2=2su5 interleaved, 8=Ean8,I=ISBN, F=farmaceutico, A=AltroIndustr., U=UPC

  '    'Pers. Alfy 06-04-21
  '    '---------------------------------------------------------
  '    'Solo per Ordini di produzione
  '    'If cbOrigine.SelectedValue = "M" AndAlso cbTipodoc.SelectedValue = "H" AndAlso cb_hhstamparifd.Enabled Then
  '    If cbOrigine.SelectedValue = "M" AndAlso cbTipodoc.SelectedValue = "H" Then
  '      'Se ho scelto la stampa diversa da "solo carta", procedo!
  '      If cb_hhstamparifd.SelectedValue <> "RPT" Then
  '        Dim dttTmpRfid As New DataTable
  '        'Carico un mio datatable con solo le righe presi dai documenti filtrati e solo quelle che hanno l'apposito campo personalizzato a "S"
  '        If oClfEtte.SelezionaDaDocMagaRFID(oApp.Ditta, cbTipodoc.SelectedValue.ToString, NTSCInt(edDaclie.Text), NTSCInt(edAclie.Text), NTSCDate(edDadata.Text), NTSCDate(edAdata.Text), NTSCInt(edDaanno.Text), NTSCInt(edAanno.Text), edDaserie.Text, edAserie.Text, NTSCInt(edDanumero.Text),
  '                               NTSCInt(edAnumero.Text), NTSCInt(edDariga.Text), NTSCInt(edAriga.Text), bDocMagazz, opColli.Checked, dttTmpRfid) Then
  '          Dim NRecordDelete As Integer = 0
  '          Dim TmpDesArt As String = String.Empty
  '          Dim TmpQtaEtic As Integer = 0
  '          Dim RisQuestDelete As DialogResult = DialogResult.Cancel
  '          'If oApp.MsgBoxInfoRedYesNo_DefNo("Cancellare tutte le etichette RFID caricate prima di oggi (" & Now.Date.ToString("dd/MM/yyyy") & ")?") = DialogResult.Yes Then oClfEtte.CancellaEtichettePrecedentiADataDBExt(Now.Date, "S", NRecordDelete)
  '          RisQuestDelete = oApp.MsgBoxInfoYesNoCancel_DefCancel("Svuotare tutte le etichette RFID caricate in precedenza?" & Chr(13) & "SI = (Tutte)" & Chr(13) & "NO = (Solo quelle caricate prima del " & Now.Date.ToString("dd/MM/yyyy") & ")" & Chr(13) & "ANNULLA = (No)")
  '          'Svuoto tutta la tabella
  '          If RisQuestDelete = DialogResult.Yes Then oClfEtte.CancellaEtichetteAllDBExt(NRecordDelete)
  '          'Cancello solo quelle precedenti alla data
  '          If RisQuestDelete = DialogResult.No Then oClfEtte.CancellaEtichettePrecedentiADataDBExt(Now.Date, "S", NRecordDelete)
  '          'Controllo che ci sono righe
  '          If dttTmpRfid.Rows.Count > 0 Then
  '            TmpDesArt = String.Empty
  '            TmpQtaEtic = 0
  '            'Ciclo le righe e cancello e inserisco riga in DB esterno con info articolo, info documento e quantità
  '            For iEticDB As Integer = 0 To dttTmpRfid.Rows.Count - 1
  '              'Lancio una delete per cancellare eventuale riga creata in precedenza con le stesse chiavi
  '              oClfEtte.CancellaEtichetteDBExt(dttTmpRfid.Rows(iEticDB)("tipork").ToString, NTSCInt(dttTmpRfid.Rows(iEticDB)("anno").ToString), dttTmpRfid.Rows(iEticDB)("serie").ToString, NTSCInt(dttTmpRfid.Rows(iEticDB)("numero").ToString), NTSCInt(dttTmpRfid.Rows(iEticDB)("riga").ToString))
  '              TmpDesArt = String.Empty
  '              TmpDesArt = oClfEtte.MyValCodiceDB(dttTmpRfid.Rows(iEticDB)("codart").ToString, "ARTICO", "S")

  '              dttTmp = Nothing
  '              If Not oCleEtte.SelectBarcode(ckBarcode.Checked, cbSolobar.SelectedValue, NTSCStr(dttTmpRfid.Rows(iEticDB)("codart").ToString), dttTmp) Then Return
  '              'Se l'articolo esiste e anche l'ean13...allora procedo.
  '              If Not String.IsNullOrEmpty(TmpDesArt) AndAlso dttTmp.Rows.Count > 0 Then
  '                TmpQtaEtic += CInt(dttTmpRfid.Rows(iEticDB)("quant").ToString)
  '                oClfEtte.InserisciEtichetteDBExt(dttTmpRfid.Rows(iEticDB)("tipork").ToString, NTSCInt(dttTmpRfid.Rows(iEticDB)("anno").ToString), dttTmpRfid.Rows(iEticDB)("serie").ToString, NTSCInt(dttTmpRfid.Rows(iEticDB)("numero").ToString), NTSCInt(dttTmpRfid.Rows(iEticDB)("riga").ToString), NTSCDate(dttTmpRfid.Rows(iEticDB)("dtdoc").ToString), dttTmp.Rows(dttTmp.Rows.Count - 1)("bc_code").ToString, TmpDesArt, CDbl(dttTmpRfid.Rows(iEticDB)("quant").ToString), TmpQtaEtic)
  '                'oClfEtte.InserisciEtichetteDBExt(dttTmpRfid.Rows(iEticDB)("tipork").ToString, NTSCInt(dttTmpRfid.Rows(iEticDB)("anno").ToString), dttTmpRfid.Rows(iEticDB)("serie").ToString, NTSCInt(dttTmpRfid.Rows(iEticDB)("numero").ToString), NTSCInt(dttTmpRfid.Rows(iEticDB)("riga").ToString), NTSCDate(dttTmpRfid.Rows(iEticDB)("dtdoc").ToString), dttTmpRfid.Rows(iEticDB)("codart").ToString, TmpDesArt, CDbl(dttTmpRfid.Rows(iEticDB)("quant").ToString), TmpQtaEtic)
  '              End If
  '            Next
  '            oApp.MsgBoxInfo("Operazione completata!" & Chr(13) & "Inviate N° " & TmpQtaEtic & " etichette RFID!" & Chr(13) & IIf(NRecordDelete > 0, "Cancellate N° " & NRecordDelete & " righe di etichette RFID!", String.Empty).ToString, True)
  '          End If
  '          'Se non ho messo "solo RFID" continuo con il resto (standard)
  '          If cb_hhstamparifd.SelectedValue = "RFI" Then
  '            oApp.MsgBoxInfo("Operazione completata!" & Chr(13) & "Inviate N° " & TmpQtaEtic & " etichette RFID!" & Chr(13) & IIf(NRecordDelete > 0, "Cancellate N° " & NRecordDelete & " righe di etichette RFID!", String.Empty).ToString)
  '            Return
  '          End If
  '        End If
  '      End If
  '    End If
  '    '---------------------------------------------------------


  '    If edNEtichettePerTC.Visible = True Then nNEtichettePerTC = NTSCInt(edNEtichettePerTC.Text)

  '    If Not oCleEtte.RiempiTmpTable(ckPrezzi.Checked, ckDatamod.Checked, NTSCDate(edDatamod.Text), ckBarcode.Checked, cbSolobar.SelectedValue,
  '                                   cbBarcode.SelectedValue, ckArtbar.Checked, cbOrigine.SelectedValue, bBD_DaCodarfo, NTSCInt(edLista.Text), opColli.Checked,
  '                                   bNEtichetteBarCode, ckValuta.Checked, NTSCInt(edListino.Text), NTSCDate(edValida.Text), NTSCInt(edValuta.Text),
  '                                   ckStampe.Checked, cbOperazione.SelectedValue, ckMisura.Checked, NTSCInt(edNumMolt.Text), bDividiQuantConfez,
  '                                   lIIstMatr, lIIstMats, bUsaTtstmatr, NTSCInt(edEtiche.Text), nBD_Listino, bModTCO,
  '                                   IIf(ckTutteLeTaglie.Checked = False, edTaglia.Text, "".PadLeft(5, "*"c)).ToString,
  '                                   lIIstMatrTmp, bStampaNegativi, dttOut, dttTmp, dttTmp2, bDocMagazz,
  '                                   bDaDoc, strCode, strCodartList, strCodart, bAvvisare, nNEtichettePerTC, ckNumeti.Checked, ckPrzNetClasseSconto.Checked,
  '                                   opNumetiDaNote.Checked) Then Return

  '    'oMenu.ResetTblInstId("TTSTMATR", bsDbAzienda, lIIstMatrTmp)
  '    '-------------------------------------------------------------------------------------------
  '    If ckTermica.Checked And nDestin = 1 Then
  '      If bDometic Then
  '        StampaDirettaDometic()
  '      Else
  '        '--- Stampo direttamente sulla stampante termica senza usare Crystal report
  '        StampaDiretta()
  '      End If
  '      Return
  '    End If
  '    '-----------------------------------------------------------------------------------------
  '    '--- determina il nome del report da stampare e la cartella relativa del registro
  '    '-----------------------------------------------------------------------------------------
  '    SettaNomeReport(strKey, strNomRpt)
  '    '-----------------------------------------------------------------------------------------
  '    '--- Lancia le stampe
  '    '-----------------------------------------------------------------------------------------
  '    nPjob = oMenu.ReportPEInit(oApp.Ditta, Me, "BSMGETTE", strKey, " ", 0, nDestin, strNomRpt, False, "Stampa etichette", False)
  '    '-----------------------------------------------------------------------------------------
  '    '--- Se la stampa è stata annullata esce
  '    '-----------------------------------------------------------------------------------------
  '    If nPjob Is Nothing Then Return
  '    '-----------------------------------------------------------------------------------------
  '    '--- Formula di selezione
  '    '-----------------------------------------------------------------------------------------
  '    If bUsaTtstmatr And (cbOrigine.SelectedValue = "M" Or cbOrigine.SelectedValue = "H") Then
  '      strCrpe = "{artico.codditt} = '" & DittaCorrente & "'" &
  '                  " And {ttstmatr.codditt} = '" & DittaCorrente & "'" &
  '                  " And {ttstmatr.instid} = " & lIIstMatr
  '    Else
  '      strCrpe = "{artico.codditt} = '" & DittaCorrente & "'" &
  '                  " And {ttstmats.codditt} = '" & DittaCorrente & "'" &
  '                  " And {ttstmats.instid} = " & lIIstMats
  '    End If

  '    '--------------------------------------------------
  '    'lancio tutti gli eventuali reports (gestisce già il multireport)
  '    For i = 1 To UBound(CType(nPjob, Array), 2)
  '      If i = 1 Then
  '        nRis = oMenu.PESetFormula(NTSCInt(CType(nPjob, Array).GetValue(0, i)), "STAMPAPREZZI", "'" & NTSCStr(IIf(ckPrezzi.Checked, "S", "N")) & "'")
  '        If ckPrezzi.Checked Then
  '          nRis = oMenu.PESetFormula(NTSCInt(CType(nPjob, Array).GetValue(0, i)), "DATAVAL", "'" & edValida.Text & "'")
  '          nRis = oMenu.PESetFormula(NTSCInt(CType(nPjob, Array).GetValue(0, i)), "LISTINO", edListino.Text)
  '          nRis = oMenu.PESetFormula(NTSCInt(CType(nPjob, Array).GetValue(0, i)), "STAMPAPREZZIVAL", "'" & NTSCStr(IIf(ckValuta.Checked, "S", "N")) & "'")
  '          If ckValuta.Checked Then
  '            nRis = oMenu.PESetFormula(NTSCInt(CType(nPjob, Array).GetValue(0, i)), "CODVALU", edValuta.Text)
  '            nRis = oMenu.PESetFormula(NTSCInt(CType(nPjob, Array).GetValue(0, i)), "DESVALU", "'" & lbDescvalu.Text & "'")
  '          End If
  '        End If
  '        nRis = oMenu.PESetSelectionFormula(NTSCInt(CType(nPjob, Array).GetValue(0, i)), strCrpe)
  '        nRis = oMenu.ReportPEVai(NTSCInt(CType(nPjob, Array).GetValue(0, i)))
  '      Else
  '        If Trim$(NTSCStr(CType(nPjob, Array).GetValue(2, i))) <> "" Then
  '          'RecordSelectionFormula
  '          nRis = oMenu.PESetSelectionFormula(NTSCInt(CType(nPjob, Array).GetValue(0, i)), CrpeResolveFormula(Me, NTSCStr(CType(nPjob, Array).GetValue(2, i)), strCrpe))
  '        End If
  '        nRis = oMenu.PESetFormula(NTSCInt(CType(nPjob, Array).GetValue(0, i)), "STAMPAPREZZI", "'" & NTSCStr(IIf(ckPrezzi.Checked, "S", "N")) & "'")
  '        If ckPrezzi.Checked Then
  '          nRis = oMenu.PESetFormula(NTSCInt(CType(nPjob, Array).GetValue(0, i)), "DATAVAL", edValida.Text)
  '          nRis = oMenu.PESetFormula(NTSCInt(CType(nPjob, Array).GetValue(0, i)), "LISTINO", edListino.Text)
  '          nRis = oMenu.PESetFormula(NTSCInt(CType(nPjob, Array).GetValue(0, i)), "STAMPAPREZZIVAL", "'" & NTSCStr(IIf(ckValuta.Checked, "S", "N")) & "'")
  '          If ckValuta.Checked Then
  '            nRis = oMenu.PESetFormula(NTSCInt(CType(nPjob, Array).GetValue(0, i)), "CODVALU", edValuta.Text)
  '            nRis = oMenu.PESetFormula(NTSCInt(CType(nPjob, Array).GetValue(0, i)), "DESVALU", "'" & lbDescvalu.Text & "'")
  '          End If
  '        End If
  '        For j = 3 To 12
  '          'FormulaJ
  '          If Trim$(NTSCStr(CType(nPjob, Array).GetValue(j, i))) <> "" Then
  '            nRis = oMenu.PESetFormula(CInt(CType(nPjob, Array).GetValue(0, i)), NTSCStr(CType(nPjob, Array).GetValue(j, i)), CrpeResolveFormula(Me, NTSCStr(CType(nPjob, Array).GetValue(j + 10, i))))
  '          End If
  '        Next j
  '        nRis = oMenu.ReportPEVai(NTSCInt(CType(nPjob, Array).GetValue(0, i)))
  '      End If
  '    Next

  '  Catch ex As Exception
  '    '-------------------------------------------------
  '    CLN__STD.GestErr(ex, Me, "")
  '    '-------------------------------------------------
  '  Finally
  '    '------------------------------------------------------------------------------------------------------------
  '    oDttgr = Nothing
  '    dttGr.Clear() : dttGr.Dispose()
  '    '------------------------------------------------------------------------------------------------------------
  '    If Not frmEtba Is Nothing Then frmEtba.Dispose()
  '    frmEtba = Nothing
  '    System.Windows.Forms.Cursor.Current = Cursors.Default
  '    lbAvviso.Text = oApp.Tr(Me, 130421093076039252, "Nessuna operazione in corso.")
  '  End Try
  'End Sub
#End Region

End Class