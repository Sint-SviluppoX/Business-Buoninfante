Imports System.Data
Imports NTSInformatica.CLN__STD
Imports System.Globalization
Imports System.Collections.Generic
Imports System
Public Class CLFVEBOLL
    Inherits CLEVEBOLL
    Public _oClhBoll As CLHVEBOLL
    Public Property oClhBoll() As CLHVEBOLL
        Get
            If _oClhBoll Is Nothing Then _oClhBoll = CType(oCldBoll, CLHVEBOLL)
            Return _oClhBoll
        End Get
        Set(ByVal value As CLHVEBOLL)
            _oClhBoll = value
        End Set
    End Property
    Public Overrides Function AfterColUpdate_CORPO_ec_quant(sender As Object, e As DataColumnChangeEventArgs) As Boolean
        Try

            'Provo ad aggiornare il campo nel piede.
            AggiornaPostiLettoCalcolati()

            Dim Tmp As Boolean = MyBase.AfterColUpdate_CORPO_ec_quant(sender, e)
            'Calcolo il volume (Pers. Alfy 25-03-21)
            'If Not IsNothing(dttArti) Then e.Row!ec_hhvol = NTSCDec(e.ProposedValue) * NTSCDec(dttArti.Rows(0)!ar_volume)
            'Calcolo peso netto e lordo (Pers. Alfy 15-03-22)
            If Not IsNothing(dttArti) Then
                'Calcolo il volume (Pers. Alfy 24-03-21)
                e.Row!ec_hhvol = NTSCDec(e.ProposedValue) * NTSCDec(dttArti.Rows(0)!ar_volume)
                'Calcolo peso netto e lordo (Pers. Alfy 15-03-22)
                e.Row!ec_hhpn = NTSCDec(e.ProposedValue) * NTSCDec(dttArti.Rows(0)!ar_pesonet)
                e.Row!ec_hhpl = NTSCDec(e.ProposedValue) * NTSCDec(dttArti.Rows(0)!ar_pesolor)
            End If
            Return Tmp
        Catch ex As Exception
            '--------------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '--------------------------------------------------------------
        End Try
    End Function
    'Ricalcola il piede da tutte le righe per non sommare due volte una riga rivisitata.
    Public Overridable Function AggiornaPostiLettoCalcolati() As Boolean
        Try
            If dsShared Is Nothing OrElse Not dsShared.Tables.Contains("CORPO") OrElse
               dttET Is Nothing OrElse dttET.Rows.Count = 0 Then Return False

            Dim corpo As DataTable = dsShared.Tables("CORPO")
            Dim campoTestata As String = "et_hhPostiLettoCalcolati"
            If Not dttET.Columns.Contains(campoTestata) Then
                campoTestata = "tm_hhPostiLettoCalcolati"
                If Not dttET.Columns.Contains(campoTestata) Then Return False
            End If

            Dim codici As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each riga As DataRow In corpo.Rows
                If riga.RowState = DataRowState.Deleted Then Continue For
                Dim codice As String = NTSCStr(riga!ec_codart).Trim()
                If codice <> "" Then codici.Add(codice)
            Next

            Dim postiPerArticolo As New Dictionary(Of String, Decimal)(StringComparer.OrdinalIgnoreCase)
            If codici.Count > 0 Then
                Dim articoli As DataTable = oClhBoll.GetPostiLettoArticoli(
                    strDittaCorrente, New List(Of String)(codici))
                If articoli Is Nothing Then Return False
                For Each articolo As DataRow In articoli.Rows
                    Dim codice As String = NTSCStr(articolo!ar_codart).Trim()
                    If codice <> "" AndAlso Not articolo.IsNull("ar_hhPostiLetto") Then
                        postiPerArticolo(codice) = NTSCDec(articolo!ar_hhPostiLetto)
                    End If
                Next
            End If

            Dim totale As Decimal = 0D
            For Each riga As DataRow In corpo.Rows
                If riga.RowState = DataRowState.Deleted Then Continue For
                Dim codice As String = NTSCStr(riga!ec_codart).Trim()
                Dim postiLetto As Decimal
                If codice <> "" AndAlso postiPerArticolo.TryGetValue(codice, postiLetto) AndAlso
                   Not riga.IsNull("ec_quant") Then
                    totale += NTSCDec(riga!ec_quant) * postiLetto
                End If
            Next
            totale = Decimal.Round(totale, 2, MidpointRounding.AwayFromZero)
            If NTSCDec(dttET.Rows(0)(campoTestata)) <> totale Then
                dttET.Rows(0)(campoTestata) = totale
            End If
            Return True
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
            Return False
        End Try
    End Function

    Public Overridable Function CalcolaTotaleVolume() As Decimal
        Try
            Dim TmpTotaleVolume As Decimal = 0
            TmpTotaleVolume = NTSCDec(dsShared.Tables("CORPO").Compute("SUM(ec_hhvol)", ""))
            Return ArrDbl(TmpTotaleVolume, oApp.NDecQta)
        Catch ex As Exception
            '--------------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '--------------------------------------------------------------	
        End Try
    End Function
    Public Overridable Function CalcolaTotalePesoNetto() As Decimal
        Try
            Dim TmpTotalePesoNetto As Decimal = 0
            If Not IsNothing(dsShared) AndAlso dsShared.Tables.Contains("CORPO") Then
                TmpTotalePesoNetto = NTSCDec(dsShared.Tables("CORPO").Compute("SUM(ec_hhpn)", ""))
            End If
            Return ArrDbl(TmpTotalePesoNetto, oApp.NDecQta)
        Catch ex As Exception
            '--------------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '--------------------------------------------------------------	
        End Try
    End Function
    Public Overridable Function CalcolaTotalePesoLordo() As Decimal
        Try
            Dim TmpTotalePesoLordo As Decimal = 0
            If Not IsNothing(dsShared) AndAlso dsShared.Tables.Contains("CORPO") Then
                TmpTotalePesoLordo = NTSCDec(dsShared.Tables("CORPO").Compute("SUM(ec_hhpl)", ""))
            End If
            Return ArrDbl(TmpTotalePesoLordo, oApp.NDecQta)
        Catch ex As Exception
            '--------------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '--------------------------------------------------------------	
        End Try
    End Function
    Public Overrides Function CalcolaTotali() As Boolean
        Try
            'dttET.Rows(0)!et_hhtotvol = CalcolaTotaleVolume()
            'Return MyBase.CalcolaTotali()
            Dim TmpTotVol As Decimal = 0
            Dim TmpTotPesoNetto As Decimal = 0
            Dim TmpTotPesoLordo As Decimal = 0
            TmpTotVol = CalcolaTotaleVolume()
            TmpTotPesoNetto = CalcolaTotalePesoNetto()
            TmpTotPesoLordo = CalcolaTotalePesoLordo()
            If TmpTotVol <> NTSCDec(dttET.Rows(0)!et_hhtotvol) Then dttET.Rows(0)!et_hhtotvol = CalcolaTotaleVolume()
            If TmpTotPesoNetto <> NTSCDec(dttET.Rows(0)!et_hhtotpn) Then dttET.Rows(0)!et_hhtotpn = CalcolaTotalePesoNetto()
            If TmpTotPesoLordo <> NTSCDec(dttET.Rows(0)!et_hhtotpl) Then dttET.Rows(0)!et_hhtotpl = CalcolaTotalePesoLordo()
            Return MyBase.CalcolaTotali()
        Catch ex As Exception
            '--------------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '--------------------------------------------------------------	
        End Try
    End Function
    Public Overrides Function SalvaDocumento(strState As String) As Boolean
        Try
            'strState: N = new, D = delete, U = update
            Dim TmpDtCorpo As New DataTable
            Dim TmpTipoDoc As String = String.Empty
            Dim TmpAnnoDoc As Integer = 0
            Dim TmpSerieDoc As String = " "
            Dim TmpNumDoc As Integer = 0
            Dim TmpRigaDoc As Integer = 0
            Dim TmpRis As Boolean = False
            Dim IsDocDaRFid As Boolean = False
            TmpTipoDoc = dttET.Rows(0)!et_tipork.ToString
            TmpAnnoDoc = NTSCInt(dttET.Rows(0)!et_anno.ToString)
            TmpSerieDoc = dttET.Rows(0)!et_serie.ToString
            TmpNumDoc = NTSCInt(dttET.Rows(0)!et_numdoc.ToString)
            'Clono il datatable del corpo prima della cancellazione per poter poi aggiornare il DB condiviso
            TmpDtCorpo = dttEC.Copy
            'Controllo se il documento proviene da RFID..se si, chiedo password per concedere modifiche e cancellazione
            If TmpTipoDoc = "T" AndAlso (strState = "D" Or strState = "U") AndAlso oClhBoll.CheckRFidDocGenerate(TmpTipoDoc, TmpAnnoDoc, TmpSerieDoc, TmpNumDoc) Then
                IsDocDaRFid = True
                Dim evnt As New NTSEventArgs(ThMsg.MSG_YESNO, oApp.Tr(Me, 130785661562704909, "Il documento risulta generato dal processo automatizzato di lettura etichette RFID!" & vbCrLf & "L'eventuale cancellazione del documento non ripristinerà le etichette utilizzate/associate!" & vbCrLf & "Sicuri di continuare?"))
                ThrowRemoteEvent(evnt)
                If evnt.RetValue = ThMsg.RETVALUE_NO Then Return False
            End If
            'Standard
            If strState <> "D" Then AggiornaPostiLettoCalcolati()
            TmpRis = MyBase.SalvaDocumento(strState)
            'Se il "SalvaDocumento" è ok e sto cancellando un carico di produzione proveniente da RFID aggiorno il DB condiviso dicendo che l'RFID usato per la generazione automatica non sarà più disponibile poichè il documento che era stato creato da esso non esiste più.
            If TmpRis AndAlso strState = "D" And TmpTipoDoc = "T" AndAlso IsDocDaRFid Then
                'Se ho cancellato un Carico da produzione indico nel DB condiviso che le etichette associate non sono più utilizzabili poichè associate ad un documento cancellato (01-06-21)
                For i As Integer = 0 To TmpDtCorpo.Rows.Count - 1
                    TmpRigaDoc = NTSCInt(TmpDtCorpo.Rows(i)("ec_riga").ToString)
                    oClhBoll.AggiornaCancellaRiferimentiDoc(TmpTipoDoc, TmpAnnoDoc, TmpSerieDoc, TmpNumDoc, TmpRigaDoc)
                Next
            End If

            '---------------------------------------------------------------------------------------------
            If Not TmpRis Then Return False
            If strState <> "D" Then
                If Not generaExcelDDT() Then oApp.MsgBoxErr("Errore nella generazione dell'Excel")
            End If
            '---------------------------------------------------------------------------------------------


            Return True
        Catch ex As Exception
            '--------------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '--------------------------------------------------------------	
        End Try
    End Function
    Public Overrides Function TestaBlocchi(bNew As Boolean) As Boolean
        Try
            Dim TmpRes As Boolean = False
            TmpRes = MyBase.TestaBlocchi(bNew)
            If (Not bNew And dsShared.Tables.Contains("TESTA")) AndAlso dsShared.Tables("TESTA").Rows.Count > 0 AndAlso dsShared.Tables("TESTA").Rows(0)("et_tipork").ToString = "T" AndAlso dsShared.Tables.Contains("CORPO") AndAlso dsShared.Tables("CORPO").Rows.Count > 0 Then
                If oClhBoll.CheckRFidDocGenerate(dsShared.Tables("TESTA").Rows(0)("et_tipork").ToString, NTSCInt(dsShared.Tables("TESTA").Rows(0)("et_anno").ToString), dsShared.Tables("TESTA").Rows(0)("et_serie").ToString, NTSCInt(dsShared.Tables("TESTA").Rows(0)("et_numdoc").ToString)) Then
                    'Alfy 03-06-21 Per il momento, se il carico di produzione è associato/è stato generato ad/da etcihette RFID, chiedo password per eventuale modifica.
                    'N.B.: Le etichette usate per la generazione del documento non saranno sbloccate...quindi riutilizzabili
                    Dim evnt As New NTSEventArgs(ThMsg.INPUTPWD, oApp.Tr(Me, 130785661562704909, "Il documento risulta generato dal processo automatizzato di lettura etichette RFID !" & vbCrLf & "Immettere password per continuare:"))
                    ThrowRemoteEvent(evnt)
                    If evnt.RetValue.Trim = "" Or (evnt.RetValue.Trim <> "" AndAlso evnt.RetValue.ToUpper <> "SINTESI") Then
                        ThrowRemoteEvent(New NTSEventArgs("", oApp.Tr(Me, 130485984222779850, "Non è stata digitata (o non è valida) la password per lo sblocco documenti. Il documento non potrà essere salvato o cancellato.")))
                        bDocNonModificabile = True
                        Return True
                    End If
                    'Dim evnt As New NTSEventArgs(ThMsg.INPUTPWD, oApp.Tr(Me, 130785661562704909, "Il documento risulta già evaso in nota di prelievo." & vbCrLf & "Immettere password per continuare:"))
                    'ThrowRemoteEvent(evnt)
                    'If evnt.RetValue.Trim = "" Or (evnt.RetValue.Trim <> "" AndAlso evnt.RetValue.ToUpper <> "SINTESI") Then
                    '  ThrowRemoteEvent(New NTSEventArgs("", oApp.Tr(Me, 130485984222779850, "Non è stata digitata (o non è valida) la password per lo sblocco documenti. Il documento non potrà essere salvato o cancellato.")))
                    '  bDocNonModificabile = True
                    '  Return True
                    'End If
                    'Dim evnt As New NTSEventArgs(ThMsg.MSG_EXCLAMATION, oApp.Tr(Me, 130785661562704909, "Il documento risulta generato dal processo automatizzato di lettura etichette RFID !" & vbCrLf & "Il documento non potrà essere salvato o cancellato."))
                    'ThrowRemoteEvent(evnt)
                    'bDocNonModificabile = True
                    Return True
                End If
            End If
            Return TmpRes
        Catch ex As Exception
            '-------------------------------------------------
            Dim strErr As String = CLN__STD.GestError(ex, Me, "", oApp.InfoError, oApp.ErrorLogFile, True)
            '-------------------------------------------------
        End Try
    End Function

    Dim strExcelDDT_Fields As String = ""
    Dim ExcelDDT_Conti As String = ""
    Dim ExcelDDT_Chiedi As Boolean = False
    Public Overrides Function LeggiRegistroDoc(strTipodoc As String) As Boolean
        Try
            If Not MyBase.LeggiRegistroDoc(strTipodoc) Then Return False

            ExcelDDT_Chiedi = oClhBoll.GetSettingBus("BSVEBOLL", "OPZIONI", ".", "ExcelDDT_Chiedi", "N", strTipodoc, "N") = "S"
            ExcelDDT_Conti = oClhBoll.GetSettingBus("BSVEBOLL", "OPZIONI", ".", "ExcelDDT_Conti", "", strTipodoc, "")
            strExcelDDT_Fields = oClhBoll.GetSettingBus("BSVEBOLL", "OPZIONI", ".", "ExcelDDT_Fields", "", strTipodoc, "")

            oClhBoll.ExcelDDT_SavePath = oClhBoll.GetSettingBus("BSVEBOLL", "OPZIONI", ".", "ExcelDDT_SavePath", "", "", "").Trim
            oClhBoll.ExcelDDT_SepDec = oClhBoll.GetSettingBus("BSVEBOLL", "OPZIONI", ".", "ExcelDDT_SepDec", "", strTipodoc, ".").Trim
            oClhBoll.ExcelDDT_SepVal = oClhBoll.GetSettingBus("BSVEBOLL", "OPZIONI", ".", "ExcelDDT_SepVal", "", strTipodoc, ";").Trim
            oClhBoll.ExcelDDT_Extension = oClhBoll.GetSettingBus("BSVEBOLL", "OPZIONI", ".", "ExcelDDT_Extension", "", "", "CSV").Trim


            Return True
        Catch ex As Exception
            '---------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '---------------------------------------------------------
        End Try
    End Function
    Public Overridable Function generaExcelDDT() As Boolean
        Dim dResult As DialogResult
        Dim dtnTmp As Dictionary(Of String, String)

        Try
            '*** Controllo se è uno dei codici conto da esportare ***
            If String.IsNullOrEmpty(ExcelDDT_Conti) Then Return True
            dResult = DialogResult.No
            For Each strConto As String In ExcelDDT_Conti.Replace(" ", "").Split(","c)
                If NTSCInt(strConto) = NTSCInt(dttET.Rows(0)!et_conto) Then
                    dResult = DialogResult.Yes
                    Exit For
                End If
            Next
            If dResult = DialogResult.No Then Return True

            '*** Chiedo se eseguire l'esportazione ***
            If ExcelDDT_Chiedi Then dResult = oApp.MsgBoxInfoYesNo_DefYes("Eseguire l'esportazione in un file excel?")

            '*** Esco se non devo esportare ***
            If dResult <> DialogResult.Yes Then Return True

            '*** Compilo il dictionary di default ***
            dtnTmp = oClhBoll.getDictionaryDDT()

            '*** Aggiungo i campi aggiuntivi al dictionary da opzione di registro ***
            For Each strExcelDDT As String In strExcelDDT_Fields.Split("|"c)
                Dim strTmp As String() = strExcelDDT.Split(";"c)
                If strTmp.Length < 2 Then Continue For

                If dtnTmp.ContainsKey(strTmp(0)) Then dtnTmp.Remove(strTmp(0))
                dtnTmp.Add(strTmp(0), strTmp(1))
            Next

            With dttET.Rows(0)
                If Not oClhBoll.generaExcelDDT(strDittaCorrente, NTSCInt(!et_numdoc), NTSCStr(!et_serie), NTSCInt(!et_anno), NTSCStr(!et_tipork)) Then Return False
            End With

            Return True
        Catch ex As Exception
            '---------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '---------------------------------------------------------
        End Try
    End Function
End Class