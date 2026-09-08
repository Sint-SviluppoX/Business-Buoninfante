Imports NTSInformatica
Imports NTSInformatica.CLN__STD
Imports System.Data
Imports System.IO
Imports System.Linq
Imports System.Runtime.InteropServices

Public Class CLE__HHID
    Inherits CLE__BASN

    Public oCldHhId As CLD__HHID

    Public strPromArtico_Impegni As String = ""
    Public strPromText_Impegni As String = ""
    Public nPromoCoVend_Impegni As Integer
    Public nPromoCodIva_Impegni As Integer

    Public strSerie_Impegni_Amazon As String = ""
    Public strSerie_Impegni_Xbio As String = ""
    Public strSerie_Impegni_SleepH As String = ""

    Public strDefArtico_Impegni As String = ""
    Public strDefPnDeiConti_Impegni As String = ""
    Public nTipoBF_Impegni As Integer
    Public nPaga_Impegni As Integer
    Public nMagaz_Impegni As Integer
    Public nListino_Impegni As Integer
    Public nCoVend_Impegni As Integer
    Public nCodIva_Impegni As Integer

    Public bPrezzoIvato As Boolean

    Public Overrides Function Init(ByRef App As CLE__APP,
                              ByRef oScriptEngine As INT__SCRIPT, ByRef oCleLbmenu As Object, ByVal strTabella As String,
                              ByVal bRemoting As Boolean, ByVal strRemoteServer As String,
                              ByVal strRemotePort As String) As Boolean
        '-----------------------------------------------------------------------------------------------------------
        If MyBase.strNomeDal = "BD__BASE" Then MyBase.strNomeDal = "BD__HHID"
        MyBase.Init(App, oScriptEngine, oCleLbmenu, strTabella, bRemoting, strRemoteServer, strRemotePort)
        oCldHhId = CType(MyBase.ocldBase, CLD__HHID)
        oCldHhId.Init(oApp)

        'oCldHhId.initFieldList()

        Return True
    End Function
    Public Overridable Function LeggiRegistroDoc() As Boolean
        Try
            '*** Leggo le varie opzioni di registro ***
            strDefPnDeiConti_Impegni = NTSCStr(oCldHhId.GetSettingBus("BS--HHID", "OPZIONI", ".", "DefPnDeiConti_Impegni", "", "", "STANDARD-AZ"))

            strPromArtico_Impegni = NTSCStr(oCldHhId.GetSettingBus("BS--HHID", "OPZIONI", ".", "PromArtico_Impegni", "", "", "D"))
            strPromText_Impegni = NTSCStr(oCldHhId.GetSettingBus("BS--HHID", "OPZIONI", ".", "PromDescr_Impegni", "", "", "Promozione articolo: {codart}"))
            nPromoCoVend_Impegni = NTSCInt(oCldHhId.GetSettingBus("BS--HHID", "OPZIONI", ".", "PromoCoVend_Impegni", "", "", "51"))
            nPromoCodIva_Impegni = NTSCInt(oCldHhId.GetSettingBus("BS--HHID", "OPZIONI", ".", "PromoCodIva_Impegni", "", "", "1022"))

            strSerie_Impegni_Amazon = NTSCStr(oCldHhId.GetSettingBus("BS--HHID", "OPZIONI", ".", "Serie_Impegni_Amazon", "", "", " "))
            strSerie_Impegni_Xbio = NTSCStr(oCldHhId.GetSettingBus("BS--HHID", "OPZIONI", ".", "Serie_Impegni_Xbio", "", "", " "))
            strSerie_Impegni_SleepH = NTSCStr(oCldHhId.GetSettingBus("BS--HHID", "OPZIONI", ".", "Serie_Impegni_SleepH", "", "", " "))

            strDefArtico_Impegni = NTSCStr(oCldHhId.GetSettingBus("BS--HHID", "OPZIONI", ".", "DefArtico_Impegni", "", "", "D"))
            nTipoBF_Impegni = NTSCInt(oCldHhId.GetSettingBus("BS--HHID", "OPZIONI", ".", "CodTipoBF_Impegni", "", "", "1"))
            nPaga_Impegni = NTSCInt(oCldHhId.GetSettingBus("BS--HHID", "OPZIONI", ".", "CodPaga_Impegni", "", "", "3"))
            nMagaz_Impegni = NTSCInt(oCldHhId.GetSettingBus("BS--HHID", "OPZIONI", ".", "CodMagaz_Impegni", "", "", "1"))
            nListino_Impegni = NTSCInt(oCldHhId.GetSettingBus("BS--HHID", "OPZIONI", ".", "CodListino_Impegni", "", "", "1"))
            nCoVend_Impegni = NTSCInt(oCldHhId.GetSettingBus("BS--HHID", "OPZIONI", ".", "CoVend_Impegni", "", "", "51"))
            nCodIva_Impegni = NTSCInt(oCldHhId.GetSettingBus("BS--HHID", "OPZIONI", ".", "CodCodIva_Impegni", "", "", "1022"))

            bPrezzoIvato = oCldHhId.GetListinoType(strDittaCorrente, nListino_Impegni)

            Return True
        Catch ex As Exception
            '---------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '---------------------------------------------------------
        End Try
    End Function

#Region "Import Dati"
    Public Overridable Function CreaTestataOrd(ByVal oImpegno As DTO_HHID_TestataImp, ByRef oCleClie As CLE__CLIE, ByRef oCleGsor As CLEORGSOR) As Boolean
        Dim dttTmp As New DataTable
        Dim strSQL As String = String.Empty
        Try
            'Ottieni controllo conto 

            '--- Ricerca del Cliente + Destinazione ----------------------------------------------------------------
            strSQL = "SELECT * FROM ANAGRA WHERE codditt = " & CStrSQL(strDittaCorrente) & " AND an_tipo = 'C' AND an_email = " & CStrSQL(oImpegno.oCliente.strEmail)
            dttTmp = ocldBase.OpenRecordset(strSQL, CLE__APP.DBTIPO.DBAZI, "ANAGRA").Tables("ANAGRA")
            If dttTmp.Rows.Count = 0 Then
                oImpegno.nConto = CreaCliente(oImpegno, oCleClie)
                'If lConto = 0 Then lConto = NTSCInt(dsShared.Tables("TABHHPS").Rows(0)!tb_hhContoPrv)
                If oImpegno.nConto = 0 Then
                    Throw New NTSException(oApp.Tr(Me, 128607611686875000, "Errore al salvataggio! Codice conto non creato! riferimento ordine: " & vbCrLf & "|" & NTSCStr(oImpegno.strRiferim) & "|"))
                    Return False
                End If
            Else
                oImpegno.nConto = CreaCliente(oImpegno, oCleClie, NTSCInt(dttTmp.Rows(0)!an_conto))
            End If

            '*** Controllo se ho trovato un codice conto ***
            If oImpegno.nConto = 0 Then
                ScriviLog($"Non è stato trovato nessun cliente legato alla mail '{NTSCStr(oImpegno.oCliente.strEmail)}'")
                Return False
            End If

            '*** Ottengo il codice destinazione ***
            oImpegno.nCoddest = GetCoddestByDescr(oImpegno, oImpegno.nConto)

            '*** Controllo se ho trovato il codice destinazione ***
            If oImpegno.nCoddest = 0 Then
                ScriviLog($"Non è stato trovato nessun indirizzo per il cliente legato all'indirizzo '{oImpegno.oCliente.strIndir}'")
                Return False
            End If
            '-------------------------------------------------------------------------------------------------------

            '*** Assegno i valori necessari alla compilazione del documento ***
            With oCleGsor.dttET.Rows(0)
                !codditt = oApp.Ditta
                !et_conto = oImpegno.nConto
                !et_coddest = oImpegno.nCoddest
                !et_tipork = oImpegno.strTipoOrd
                !et_anno = oImpegno.nAnnoOrd
                !et_serie = oImpegno.strSerieOrd
                !et_numdoc = oImpegno.lNumTmpOrd
                !et_scopag = oImpegno.nScoPag
                !et_datdoc = oImpegno.dDatDoc
                !et_riferim = oImpegno.strRiferim.Trim
                !et_scorpo = oImpegno.strScorporo

                !et_tipobf = oImpegno.nTipoBf
                !et_magaz = oImpegno.nCodMaga
                !et_listino = oImpegno.nListino
                !et_codpaga = oImpegno.nCodPaga

                !et_speacc = oImpegno.nSpeAcc

                !et_hhdtImpo = oImpegno.dDtaImpo
                !et_hhorImpo = oImpegno.dOrImpo
            End With

            '*** Controllo se ho compilato tutti i dati necessari per la testata ***
            If Not oCleGsor.OkTestata Then
                ScriviLog("Ci sono errori nella testata del documento")
                Return False
            End If

            Return True
        Catch ex As Exception
            ScriviLog("Errore durante la creazione della testata" & vbCrLf & ex.Message)
            CLN__STD.GestErr(ex, Me, "")
        End Try
    End Function
    Public Overridable Function CreaRigheOrd(ByVal oImpegno As DTO_HHID_TestataImp, oCleGsor As CLEORGSOR, ByRef Optional bArticoloDef As Boolean = False) As Boolean
        Dim strSQL As String = String.Empty
        Dim dttTmp As DataTable
        Try
            '*** Ciclo tutte le righe ***
            For Each oRiga As DTO_HHID_CorpoImp In oImpegno.oCorpo
                If String.IsNullOrEmpty(oRiga.strCodArt) Then
                    ScriviLog($"Nessuno legato al codice '{oRiga.strCodAlt}' Verra inserito come articolo '{strDefArtico_Impegni}'")
                    oRiga.strCodArt = strDefArtico_Impegni
                    bArticoloDef = True
                ElseIf Not oCldHhId.ValCodiceDb(oRiga.strCodArt, strDittaCorrente, "ARTICO", "S") Then
                    ScriviLog($"Codice articolo '{oRiga.strCodArt}' non trovato. Verra inserito come articolo '{strDefArtico_Impegni}'")
                    oRiga.strCodArt = strDefArtico_Impegni
                End If

                '*** Aggiungo la riga dell'articolo al corpo ***
                If Not oCleGsor.AggiungiRigaCorpo(False, oRiga.strCodArt, 0, 0) Then Return False

                '*** Compilo gli altri dati del rigo ***
                With oCleGsor.dttEC.Rows(oCleGsor.dttEC.Rows.Count - 1)
                    !ec_descr = NTSCStr(oRiga.strDescr)
                    !ec_quant = oRiga.dQuant

                    If bPrezzoIvato Then !ec_preziva = oRiga.dPrezzoIva Else !ec_prezzo = oRiga.dPrezzo

                    !ec_controp = oRiga.nControp
                    !ec_codiva = oRiga.nCodIva

                    !ec_colpre = oRiga.dColPre
                    !ec_quapre = oRiga.dQuaPre

                    '*** Ottengo il numero riga che mi servirà in seguito per gli articoli kit ***
                    oRiga.nRiga = NTSCInt(!ec_riga)
                End With

                '*** Salvataggio del rigo ***
                If Not oCleGsor.RecordSalva(oCleGsor.dttEC.Rows.Count - 1, False, Nothing) Then
                    '*** Scrivo il log ***
                    ScriviLog("Errore nel salvataggio della riga del documento")

                    '*** Canello la riga errata ***
                    oCleGsor.dttEC.Rows(oCleGsor.dttEC.Rows.Count - 1).Delete()
                    Return False
                End If

                If oRiga.dPrzPromo <> 0 Then
                    If Not oCleGsor.AggiungiRigaCorpo(False, strPromArtico_Impegni, 0, 0) Then Return False
                    With oCleGsor.dttEC.Rows(oCleGsor.dttEC.Rows.Count - 1)
                        If oRiga.strTipoPromo = "ART" Then
                            !ec_descr = NTSCStr(strPromText_Impegni.Replace("{codart}", oRiga.strCodArt))
                        ElseIf oRiga.strTipoPromo = "TOT" Then
                            '!ec_descr = NTSCStr(strPromText_Impegni)
                        Else
                            '!ec_descr = NTSCStr(strPromText_Impegni)
                        End If

                        !ec_quant = 1

                        If bPrezzoIvato Then !ec_preziva = oRiga.dPrzPromo Else !ec_prezzo = oRiga.dPrzPromo

                        !ec_controp = oRiga.nContropPromo
                        !ec_codiva = oRiga.nCodIvaPromo
                    End With

                    '*** Salvataggio del rigo ***
                    If Not oCleGsor.RecordSalva(oCleGsor.dttEC.Rows.Count - 1, False, Nothing) Then
                        '*** Scrivo il log ***
                        ScriviLog("Errore nel salvataggio della riga di promozione del documento")

                        '*** Canello la riga errata ***
                        oCleGsor.dttEC.Rows(oCleGsor.dttEC.Rows.Count - 1).Delete()
                        Return False
                    End If
                End If

            Next

            Return True
        Catch ex As Exception
            '---------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '---------------------------------------------------------
        End Try
    End Function
    Public Overridable Function SettaPiedeOrd(ByVal oImpegno As DTO_HHID_TestataImp, oCleGsor As CLEORGSOR) As Boolean
        Try
            '*** Calcolo i totali del piede ***
            oCleGsor.CalcolaTotali()
            Return True
        Catch ex As Exception
            '---------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '---------------------------------------------------------
        End Try
    End Function

    '--- Creazione di un nuovo cliente ------------------------------------------------------------------
    Public nCodiceMastro As Integer = 401 'codice mastro cli/for
    Public strTipoCliFor As String = "C" ' "F" per fornitori
    Public strTipoNumerazCliFor As String = "CC" ' "FF" per fornitori
    Public Overridable Function CreaCliente(ByVal oImpegno As DTO_HHID_TestataImp, oCleClie As CLE__CLIE, Optional ByVal lContoTmp As Integer = -1) As Integer

        Try
            Dim strDittaCorrente As String = oApp.Ditta

            Dim bNew As Boolean = True
            Dim dsClie As New DataSet
            Dim strApriWhere As String = ""

            nCodiceMastro = GetCodMastroClienti()
            If nCodiceMastro = 0 Then Return 0

            'initext
            oCleClie.bServer = CBool(IIf(UCase(oApp.Profilo) <> "BUSOFFLINE", True, False))
            oCleClie.strTestSalvaCfPiva = CType(oApp.oMenu, CLE__MENU).GetSettingBus("BS--CLIE", "OPZIONI", ".", "TestSalvaCfPiva", "N", " ", "N").ToString
            oCleClie.bNonProporreSiglaRic = CBool(CType(oApp.oMenu, CLE__MENU).GetSettingBus("BS--CLIE", "OPZIONI", ".", "NonProporreSiglaRic", "0", " ", "0"))
            oCleClie.strGeneraIdPswClienti = CType(oApp.oMenu, CLE__MENU).GetSettingBus("BS--CLIE", "OPZIONI", ".", "GeneraIdPswClienti", "0", " ", "0").ToString
            oCleClie.strAnagenDtIniz = CType(oApp.oMenu, CLE__MENU).GetSettingBus("BS--CLIE", "OPZIONI", ".", "AnagenDtIniz", "0", " ", "0")
            oCleClie.bGesttabcont = CBool(CType(oApp.oMenu, CLE__MENU).GetSettingBus("OPZIONI", ".", ".", "GestTabcont", IIf(CBool(CType(oApp.oMenu, CLE__MENU).ModuliDittaDitt(strDittaCorrente) And bsModPM), "-1", "0").ToString, " ", IIf(CBool(CType(oApp.oMenu, CLE__MENU).ModuliDittaDitt(strDittaCorrente) And bsModPM), "-1", "0").ToString))
            oCleClie.strBloccaavvertifido = CType(oApp.oMenu, CLE__MENU).GetSettingBus("BSCGDCST", "OPZIONI", ".", "Bl_cliente_sup_fido", " ", " ", " ")
            oCleClie.strBloccainsolu = CType(oApp.oMenu, CLE__MENU).GetSettingBus("BSCGDCST", "OPZIONI", ".", "Bl_cliente_per_insol", " ", " ", " ")
            oCleClie.strBloccaRDScadute = CType(oApp.oMenu, CLE__MENU).GetSettingBus("BSCGDCST", "OPZIONI", ".", "Bl_cliente_RD_scadute", "N", " ", "N")
            oCleClie.bRiapriSuSalva = CBool(CType(oApp.oMenu, CLE__MENU).GetSettingBus("BS--CLIE", "OPZIONI", ".", "RiapriSuSalva", "0", " ", "0"))
            oCleClie.nCodpagaInAddNew = CInt(CType(oApp.oMenu, CLE__MENU).GetSettingBus("BS--CLIE", "OPZIONI", ".", "Pagamento_Nuovo_Cliente", "0", " ", "0"))
            oCleClie.nListinoInAddNew = CInt(CType(oApp.oMenu, CLE__MENU).GetSettingBus("BS--CLIE", "OPZIONI", ".", "Listino_Nuovo_Cliente", "1", " ", "1"))
            oCleClie.strDefaultUserCrm = CType(oApp.oMenu, CLE__MENU).GetSettingBus("BS--CLIE", "OPZIONI", ".", "UtentePredefinitoCrm", "Admin", " ", "Admin")
            If oCleClie.nListinoInAddNew = 0 Then oCleClie.nListinoInAddNew = 1
            oCleClie.strPasswbl = CType(oApp.oMenu, CLE__MENU).GetSettingBus("BS--CLIE", "OPZIONI", ".", "Passwbl", "nts", " ", "nts")
            oCleClie.bGestAlert = CBool(Val(CType(oApp.oMenu, CLE__MENU).GetSettingBus("BS--CLIE", "OPZIONI", ".", "Abilita_Alert", "0", " ", "0")))
            oCleClie.strVoceFinClie = CType(oApp.oMenu, CLE__MENU).GetSettingBus("BS--CLIE", "OPZIONI", ".", "VoceFinanziariaCliente", "", ".", "")
            oCleClie.strVoceFinForn = CType(oApp.oMenu, CLE__MENU).GetSettingBus("BS--CLIE", "OPZIONI", ".", "VoceFinanziariaForn", "", ".", "")

            '--------------------------------------------------------------------------------------------------------------
            '--- Lettura dell'instid su TTDESTDIV di appoggio, per la stampa degli Indirizzi dalla modale relativa
            '--------------------------------------------------------------------------------------------------------------
            oCleClie.lIITtdestdiv = CType(oApp.oMenu, CLE__MENU).GetTblInstId("TTDESTDIV", False)

            'NB:---------------------------------------------------------------------------------------
            oCleClie.strTipoConto = strTipoCliFor
            'NB:---------------------------------------------------------------------------------------

            'controllo iniziale
            If Not oCleClie.Apri(strDittaCorrente, False, lContoTmp, "", dsClie) Then
                Return 0
            End If
            dsClie.AcceptChanges()

            Dim dttPeca As New DataTable
            '------------------------------------------------
            'CRM: se l'operatore non è stato codificato e non ha un ruolo non può operare
            oCleClie.bModuloAS = False
            oCleClie.bModuloCRM = False
            If CBool(CType(oApp.oMenu, CLE__MENU).ModuliDittaDitt(strDittaCorrente) And CLN__STD.bsModAS) Then oCleClie.bModuloAS = True
            If CBool(CType(oApp.oMenu, CLE__MENU).ModuliExtDittaDitt(strDittaCorrente) And CLN__STD.bsModExtCRM) Then oCleClie.bModuloCRM = True
            If CBool(CType(oApp.oMenu, CLE__MENU).ModuliSupDittaDitt(strDittaCorrente) And CLN__STD.bsModSupWCR) Then oCleClie.bModuloCRM = True
            If oCleClie.bModuloCRM Then
                oCleClie.bIsCRMUser = CType(oApp.oMenu, CLE__MENU).IsCrmUser(strDittaCorrente, oCleClie.bAmm, oCleClie.strAccvis, oCleClie.strAccmod, oCleClie.strRegvis, oCleClie.strRegmod)

                If oCleClie.bIsCRMUser Then
                    oCleClie.lCodorgaOperat = CType(oApp.oMenu, CLE__MENU).RitornaCodorgaDaOpnome(strDittaCorrente, oCleClie.nCodcageoperat)
                    If oCleClie.lCodorgaOperat = 0 Then
                        oApp.MsgBoxErr(oApp.Tr(Me, 127791222142500000, "Attenzione!" & vbCrLf & "L'operatore '|" & oApp.User.Nome &
                         "|' (CRM) non è associato all'organizzazione della ditta corrente '|" & strDittaCorrente & "|'." & vbCrLf &
                         "Impossibile continuare."))
                        Return 0
                    End If
                End If
            Else
                If oCleClie.bModuloAS Then oCleClie.lCodorgaOperat = CType(oApp.oMenu, CLE__MENU).RitornaCodorgaDaOpnome(strDittaCorrente, oCleClie.nCodcageoperat)
            End If    ' If bModuloCRM Then

            If Not oCleClie.LeggiDatiDitta(strDittaCorrente, CBool(CType(oApp.oMenu, CLE__MENU).GetSettingBus("BS--CLIE", "OPZIONI", ".", "GestAnaExt", "0", " ", "0"))) Then
                Return 0
            End If

            If CBool(CType(oApp.oMenu, CLE__MENU).ModuliSupDittaDitt(strDittaCorrente) And bsModSupCAE) Then
                CType(oApp.oMenu, CLE__MENU).ValCodiceDb("1", strDittaCorrente, "TABPECA", "N", , dttPeca)
                If dttPeca.Rows.Count = 0 Then
                    oApp.MsgBoxErr(oApp.Tr(Me, 129319677734956014, "Tabella delle Personalizzazioni CA-DC (Globale) non configurata. Imposibile continuare"))
                    Return 0
                End If
                oCleClie.bCampiCAEAttivi = CBool(IIf(NTSCStr(dttPeca.Rows(0)!tb_richcli) = "S", True, False))
            End If
            'fine initext


            Dim lContoProgrMoltip As Integer = 0
            Dim lProgr As Integer = 0

            Dim dttTmp As New DataTable
            Dim dttAnaz As New DataTable
            CType(oApp.oMenu, CLE__MENU).ValCodiceDb(strDittaCorrente, strDittaCorrente, "TABANAZ", "S", "", dttAnaz)
            CType(oApp.oMenu, CLE__MENU).ValCodiceDb(dttAnaz.Rows(0)!tb_azcodpcon.ToString, "", "TABPCON", "S", "", dttTmp)
            Select Case dttTmp.Rows(0)!tb_struttura.ToString
                Case "A" : lContoProgrMoltip = 100000
                Case "B" : lContoProgrMoltip = 1000000
                Case "C" : lContoProgrMoltip = 1000000
                Case "D" : lContoProgrMoltip = 100000
                Case "S" : lContoProgrMoltip = 10000
            End Select

            'leggo numerazione mastro cli/for
            lProgr = oCleClie.oCldClie.LegNuma(strDittaCorrente, strTipoNumerazCliFor, "", nCodiceMastro, True)
            'aggiorno la numerazione mastro cli/for
            Dim strMsg As String = ""
            lProgr = oCleClie.oCldClie.AggNuma(strDittaCorrente, strTipoNumerazCliFor, "", nCodiceMastro, lProgr, True, True, strMsg)
            If strMsg <> "" Then
                MsgBox(strMsg)
            End If

            If lContoTmp = -1 Then
                Dim dttAnag As New DataTable
                dttAnag = ocldBase.OpenRecordset("SELECT * FROM anagra WHERE codditt = " & CStrSQL(strDittaCorrente) & " AND an_tipo = 'C' AND an_email = " & CStrSQL(oImpegno.oCliente.strEmail), CLE__APP.DBTIPO.DBAZI, "ANAGRA").Tables("ANAGRA")

                If dttAnag IsNot Nothing AndAlso dttAnag.Rows.Count > 0 Then
                    lContoTmp = NTSCInt(dttAnag.Rows(0)!an_conto)
                    bNew = False
                Else
                    'lcontotmp= (nCodiceMastro * lContoProgrMoltip) + lProgr
                    GetCodMastroClienti(lContoTmp)
                    bNew = True
                End If
            Else
                If ocldBase.ValCodiceDb(NTSCStr(lContoTmp), strDittaCorrente, "ANAGRA", "N") Then bNew = False Else bNew = True
            End If

            If Not oCleClie.Apri(strDittaCorrente, bNew, lContoTmp, strApriWhere, dsClie) Then
                oApp.MsgBoxErr("Impossbile creare il conto!")
                Return 0
            End If

            If Not oCleClie.CaricaDestdiv(dsClie.Tables("ANAGRA").Rows(0)) Then Return 0
            If Not oCleClie.CaricaAnaext(dsClie.Tables("ANAGRA").Rows(0), 0) Then Return 0
            If Not oCleClie.CaricaTabelleCollegate(dsClie.Tables("ANAGRA").Rows(0)) Then Return 0
            Dim bContoMovimentato As Boolean = False
            oCleClie.CaricaColonneUnbound(dsClie.Tables("ANAGRA").Rows(0), bContoMovimentato)
            oCleClie.bHasChanges = True
            '--- Lead --------------------------------------------------------------------------
            If (oCleClie.bModuloCRM Or oCleClie.bModuloAS) Then
                oCleClie.lLead = oCleClie.CercaLeadDaConto(NTSCInt(lContoTmp), 0,
                  oCleClie.ModuliDittaDitt(strDittaCorrente), oCleClie.ModuliExtDittaDitt(strDittaCorrente),
                  oCleClie.ModuliSupDittaDitt(strDittaCorrente))
                oCleClie.lObiettivo = oCleClie.GetObiettivoDaLead(oCleClie.lLead)
                oCleClie.NuovoAnagra(lContoTmp, nCodiceMastro, 0, oCleClie.lLead, "", dsClie.Tables("ANAGRA").Rows(0))
            Else
                oCleClie.lLead = 0
                oCleClie.lObiettivo = 0
            End If

            '--- Imposto i campi dell'anagrafica --------------------------------------------------------------------------------
            With dsClie.Tables("ANAGRA").Rows(0)
                If bNew Then
                    !an_codmast = nCodiceMastro 'questo è obbligatorio
                    !an_codpcon = strDefPnDeiConti_Impegni
                End If
                !an_descr1 = oImpegno.oCliente.strNome

                If oCleClie.IsEmail(NTSCStr(oImpegno.oCliente.strEmail).Trim) Then
                    !an_email = NTSCStr(oImpegno.oCliente.strEmail).Trim
                Else
                    Throw New NTSException(oApp.Tr(Me, 128607611686875000, "Email '" & NTSCStr(oImpegno.oCliente.strEmail) & "' non valida per il cliente: " & NTSCStr(oImpegno.oCliente.strNome) & "!"))
                    Return 0
                End If

                Dim strTelef As String = NTSCStr(oImpegno.oCliente.strPhone)
                If strTelef <> "0000000000" AndAlso Not String.IsNullOrEmpty(strTelef.Trim) Then
                    If strTelef.Length > 18 Then !an_telef = strTelef.Substring(0, 18) Else !an_telef = strTelef
                Else
                    !an_telef = ""
                End If
                '----------------------------------------------------------------------------------------------------------------
                If bNew Then
                    '*** Imposto il cliente come privato ***
                    !an_privato = "S"

                    '*** Salvo l'indirizzo del clietne ***
                    !an_indir = oImpegno.oCliente.strIndir

                    '*** Salvo gli altri dati inerenti all'indirizzo ***
                    !an_citta = oImpegno.oCliente.strCitta
                    '----------------------------------------------------------
                    !an_codcomu = oImpegno.oCliente.strCodComune
                    !an_prov = oImpegno.oCliente.strCodProv
                    !an_stato = oImpegno.oCliente.strCodStato
                    !an_cap = oImpegno.oCliente.strCap
                    '----------------------------------------------------------
                End If
            End With

            'salva
            If Not oCleClie.Salva(False) Then Return 0
            '--------------------------------------------------------------------------------------------------------
            If Not ImpostaDestinazioniDiverse(oImpegno, dsClie, oCleClie, NTSCInt(dsClie.Tables("ANAGRA").Rows(0)!an_conto)) Then Return 0
            If Not oCleClie.Salva(False) Then Return 0
            '--------------------------------------------------------------------------------------------------------
            Return NTSCInt(dsClie.Tables("ANAGRA").Rows(0)!an_conto)

        Catch ex As Exception
            '---------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '---------------------------------------------------------
        End Try
    End Function
    Public Overridable Function ImpostaDestinazioniDiverse(ByVal oImpegno As DTO_HHID_TestataImp, ByVal dsClie As DataSet, ByVal oCleClie As CLE__CLIE, ByVal lContoTmp As Integer) As Boolean
        Dim StrWhere As String = String.Empty
        Dim StrSQL As String = String.Empty
        Dim i As Integer = 0
        Dim bNew As Boolean = False

        Try
            '------------------------------------------------------------------
            oCleClie.DesgSetDataTable(strDittaCorrente, dsClie.Tables("DESTDIV"))
            '-------------------------------------------------------------------

            '*** Controllo se l'indirizzo combacia ***
            StrWhere = $"dd_inddest = {CStrSQL(oImpegno.oCliente.strIndir)}"
            If Not String.IsNullOrEmpty(oImpegno.oCliente.strCodComune) Then StrWhere &= $" AND dd_codcomu = {CStrSQL(oImpegno.oCliente.strCodComune)}"
            If Not String.IsNullOrEmpty(oImpegno.oCliente.strCodStato) Then StrWhere &= $" AND dd_prodest = {CStrSQL(oImpegno.oCliente.strCodStato)}"
            If Not String.IsNullOrEmpty(oImpegno.oCliente.strCodNazione) Then StrWhere &= $" AND dd_stato = {CStrSQL(oImpegno.oCliente.strCodNazione)}"
            Dim dtrDest As DataRow() = dsClie.Tables("DESTDIV").Select(StrWhere)

            '*** Imposto la destinaizone esistente o nuova ***
            If dtrDest.Length > 0 Then
                i = dsClie.Tables("DESTDIV").Rows.IndexOf(dtrDest(0))
                bNew = False
            Else
                oCleClie.DesgNuovo()
                i = dsClie.Tables("DESTDIV").Rows.Count - 1
                bNew = True
            End If

            With dsClie.Tables("DESTDIV").Rows(i)
                If NTSCStr(oImpegno.oCliente.strNome).Length <= 40 Then !dd_nomdest = oImpegno.oCliente.strNome Else !dd_nomdest = NTSCStr(oImpegno.oCliente.strNome).Substring(0, 40)

                !dd_inddest = oImpegno.oCliente.strIndir
                !dd_note = oImpegno.oCliente._strIndir

                !dd_locdest = oImpegno.oCliente.strCitta

                Dim strTelef As String = oImpegno.oCliente.strPhone
                If strTelef <> "0000000000" AndAlso Not String.IsNullOrEmpty(strTelef.Trim) Then
                    If strTelef.Length > 18 Then !dd_telef = strTelef.Substring(0, 18) Else !dd_telef = strTelef
                Else
                    !dd_telef = ""
                End If

                '----------------------------------------------------------
                !dd_codcomu = oImpegno.oCliente.strCodComune
                !dd_prodest = oImpegno.oCliente.strCodStato
                !dd_stato = oImpegno.oCliente.strCodNazione
                !dd_capdest = oImpegno.oCliente.strCap
                '----------------------------------------------------------
            End With
            Return True
        Catch ex As Exception
            '---------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '---------------------------------------------------------
        End Try
    End Function
#End Region

#Region "Gestione"
    Public Overridable Function getImpegniExcelAmazon(ByVal strPath As String, ByVal strFoglio As String, ByRef Optional dsOut As DataSet = Nothing) As Boolean
        Dim strFields As String
        Dim strTipo As String = "Amazon"
        Try
            '*** Inizializzo il dictionary per rinominare le colonne dell' ***
            oCldHhId.initFieldList(strTipo)
            '*** Compongo la select, per selezionare solo i dati che mi interessano ***
            strFields = oCldHhId.FieldsList.Item("ORDER-ID")
            strFields &= ", " & oCldHhId.FieldsList.Item("PURCHASE-DATE")
            strFields &= ", " & oCldHhId.FieldsList.Item("BUYER-EMAIL")
            strFields &= ", " & oCldHhId.FieldsList.Item("BUYER-NAME")
            strFields &= ", " & oCldHhId.FieldsList.Item("BUYER-PHONE-NUMBER")
            strFields &= ", " & oCldHhId.FieldsList.Item("SHIP-ADDRESS-1")
            strFields &= ", " & oCldHhId.FieldsList.Item("SHIP-ADDRESS-2")
            strFields &= ", " & oCldHhId.FieldsList.Item("SHIP-ADDRESS-3")
            strFields &= ", " & oCldHhId.FieldsList.Item("SHIP-CITY")
            strFields &= ", " & oCldHhId.FieldsList.Item("SHIP-STATE")
            strFields &= ", " & oCldHhId.FieldsList.Item("SHIP-POSTAL-CODE")
            strFields &= ", " & oCldHhId.FieldsList.Item("SHIP-COUNTRY")
            strFields &= ", " & oCldHhId.FieldsList.Item("SHIP-PHONE-NUMBER")
            strFields &= ", " & oCldHhId.FieldsList.Item("SHIPPING-PRICE")
            strFields &= ", " & oCldHhId.FieldsList.Item("SHIPPING-TAX")


            '*** Leggo i dati dal testo ***
            'oCldHhId.GetDataFromCSV(strPath.Replace(strPath.Split("\"c).Last, ""), strPath.Split("\"c).Last, "TESTA", "*", "", "", "", "", dsOut)
            oCldHhId.LetturaExcel(strPath, "Testa", strTipo, strFields, "", strFields, oCldHhId.FieldsList.Item("PURCHASE-DATE"), "", dsOut)

            '*** Compongo la select, per selezionare solo i dati che mi interessano ***
            strFields = oCldHhId.FieldsList.Item("ORDER-ID")
            strFields &= ", " & oCldHhId.FieldsList.Item("PURCHASE-DATE")
            strFields &= ", " & oCldHhId.FieldsList.Item("BUYER-EMAIL")
            strFields &= ", " & oCldHhId.FieldsList.Item("BUYER-NAME")
            strFields &= ", " & oCldHhId.FieldsList.Item("BUYER-PHONE-NUMBER")
            strFields &= ", " & oCldHhId.FieldsList.Item("SHIP-ADDRESS-1")
            strFields &= ", " & oCldHhId.FieldsList.Item("SHIP-ADDRESS-2")
            strFields &= ", " & oCldHhId.FieldsList.Item("SHIP-ADDRESS-3")
            strFields &= ", " & oCldHhId.FieldsList.Item("SHIP-CITY")
            strFields &= ", " & oCldHhId.FieldsList.Item("SHIP-STATE")
            strFields &= ", " & oCldHhId.FieldsList.Item("SHIP-POSTAL-CODE")
            strFields &= ", " & oCldHhId.FieldsList.Item("SHIP-COUNTRY")
            strFields &= ", " & oCldHhId.FieldsList.Item("SHIP-PHONE-NUMBER")
            strFields &= ", " & oCldHhId.FieldsList.Item("SHIPPING-PRICE")

            strFields &= ", " & oCldHhId.FieldsList.Item("ORDER-ITEM-ID")
            strFields &= ", " & oCldHhId.FieldsList.Item("SKU")
            strFields &= ", " & oCldHhId.FieldsList.Item("PRODUCT-NAME")
            strFields &= ", " & oCldHhId.FieldsList.Item("QUANTITY-PURCHASED")
            strFields &= ", " & oCldHhId.FieldsList.Item("ITEM-PRICE")
            strFields &= ", " & oCldHhId.FieldsList.Item("ITEM-TAX")
            strFields &= ", " & oCldHhId.FieldsList.Item("ITEM-PROMOTION-DISCOUNT")

            '*** Leggo i dati dal testo ***
            'oCldHhId.LetturaTesto(strPath, "Corpo", strFields, Nothing, Nothing, dsOut)
            oCldHhId.LetturaExcel(strPath, "Corpo", strTipo, strFields, "", strFields, "", "", dsOut)

            Return True
        Catch ex As Exception
            '---------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            Return False
            '---------------------------------------------------------
        End Try
    End Function
    Public Overridable Function getImpegniExcelXbio(ByVal strPath As String, ByVal strFoglio As String, ByRef Optional dsOut As DataSet = Nothing) As Boolean
        Dim strFields As String
        Dim strTipo As String = "Xbio"
        Try
            '*** Inizializzo il dictionary per rinominare le colonne dell' ***
            oCldHhId.initFieldList(strTipo)
            '*** Compongo la select, per selezionare solo i dati che mi interessano ***
            strFields = oCldHhId.FieldsList.Item("ORDER-REFERENCE")
            strFields &= ", " & oCldHhId.FieldsList.Item("ORDER-CREATION-DATE")
            strFields &= ", " & oCldHhId.FieldsList.Item("CUSTOMER-EMAIL")
            strFields &= ", " & oCldHhId.FieldsList.Item("INVOICE-ADDRESS-FIRSTNAME")
            strFields &= ", " & oCldHhId.FieldsList.Item("INVOICE-ADDRESS-LASTNAME")
            strFields &= ", " & oCldHhId.FieldsList.Item("DELIVERY-ADDRESS-PHONE")
            strFields &= ", " & oCldHhId.FieldsList.Item("DELIVERY-ADDRESS-DNI")
            strFields &= ", " & oCldHhId.FieldsList.Item("DELIVERY-ADDRESS1")
            strFields &= ", " & oCldHhId.FieldsList.Item("DELIVERY-ADDRESS2")
            strFields &= ", " & oCldHhId.FieldsList.Item("DELIVERY-ADDRESS-POSTCODE")
            strFields &= ", " & oCldHhId.FieldsList.Item("DELIVERY-ADDRESS-CITY")
            strFields &= ", " & oCldHhId.FieldsList.Item("DELIVERY-ADDRESS-STATE-NAME")
            strFields &= ", " & oCldHhId.FieldsList.Item("DELIVERY-ADDRESS-COUNTRY-NAME")
            strFields &= ", " & oCldHhId.FieldsList.Item("TOTAL-SHIPPING-TAX-EXCLUDED")
            strFields &= ", " & oCldHhId.FieldsList.Item("TOTAL-SHIPPING-TAX-INCLUDED")

            '*** Leggo i dati dal testo ***
            'oCldHhId.GetDataFromCSV(strPath.Replace(strPath.Split("\"c).Last, ""), strPath.Split("\"c).Last, "TESTA", "*", "", "", "", "", dsOut)
            oCldHhId.LetturaExcel(strPath, "Testa", strTipo, strFields, "", strFields, oCldHhId.FieldsList.Item("ORDER-CREATION-DATE"), "", dsOut)

            '*** Compongo la select, per selezionare solo i dati che mi interessano ***
            strFields = oCldHhId.FieldsList.Item("ORDER-REFERENCE")
            strFields &= ", " & oCldHhId.FieldsList.Item("ORDER-CREATION-DATE")
            strFields &= ", " & oCldHhId.FieldsList.Item("CUSTOMER-EMAIL")
            strFields &= ", " & oCldHhId.FieldsList.Item("INVOICE-ADDRESS-FIRSTNAME")
            strFields &= ", " & oCldHhId.FieldsList.Item("INVOICE-ADDRESS-LASTNAME")
            strFields &= ", " & oCldHhId.FieldsList.Item("DELIVERY-ADDRESS-PHONE")
            strFields &= ", " & oCldHhId.FieldsList.Item("DELIVERY-ADDRESS-DNI")
            strFields &= ", " & oCldHhId.FieldsList.Item("DELIVERY-ADDRESS1")
            strFields &= ", " & oCldHhId.FieldsList.Item("DELIVERY-ADDRESS2")
            strFields &= ", " & oCldHhId.FieldsList.Item("DELIVERY-ADDRESS-POSTCODE")
            strFields &= ", " & oCldHhId.FieldsList.Item("DELIVERY-ADDRESS-CITY")
            strFields &= ", " & oCldHhId.FieldsList.Item("DELIVERY-ADDRESS-STATE-NAME")
            strFields &= ", " & oCldHhId.FieldsList.Item("DELIVERY-ADDRESS-COUNTRY-NAME")
            strFields &= ", " & oCldHhId.FieldsList.Item("TOTAL-SHIPPING-TAX-EXCLUDED")
            strFields &= ", " & oCldHhId.FieldsList.Item("TOTAL-SHIPPING-TAX-INCLUDED")

            strFields &= ", " & oCldHhId.FieldsList.Item("COMBINATION-EAN13")
            strFields &= ", " & oCldHhId.FieldsList.Item("RIFERIMENTO-DEL-PRODOTTO")
            strFields &= ", " & oCldHhId.FieldsList.Item("PRODUCT-NAME-WITH-COMBINATION")
            strFields &= ", " & oCldHhId.FieldsList.Item("PRODUCT-QUANTITY")
            strFields &= ", " & oCldHhId.FieldsList.Item("PRODUCT-UNIT-PRICE-TAX-INCLUDED")
            strFields &= ", " & oCldHhId.FieldsList.Item("PRODUCT-UNIT-PRICE-TAX-EXCLUDED")
            strFields &= ", " & oCldHhId.FieldsList.Item("TOTAL-DISCOUNTS-TAX-INCLUDED")
            strFields &= ", " & oCldHhId.FieldsList.Item("TOTAL-DISCOUNTS-TAX-EXCLUDED")

            strFields &= ", " & oCldHhId.FieldsList.Item("INVOICETAXTYPE-TAXID-TAXNAME-TAXAMOUNT")


            '*** Leggo i dati dal testo ***
            'oCldHhId.LetturaTesto(strPath, "Corpo", strFields, Nothing, Nothing, dsOut)
            oCldHhId.LetturaExcel(strPath, "Corpo", strTipo, strFields, "", strFields, "", "", dsOut)

            Return True
        Catch ex As Exception
            '---------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            Return False
            '---------------------------------------------------------
        End Try
    End Function
    Public Overridable Function getImpegniExcelSleepH(ByVal strPath As String, ByVal strFoglio As String, ByRef Optional dsOut As DataSet = Nothing) As Boolean
        Dim strFields As String
        Dim strTipo As String = "SleepH"
        Try
            '*** Inizializzo il dictionary per rinominare le colonne dell' ***
            oCldHhId.initFieldList(strTipo)
            '*** Compongo la select, per selezionare solo i dati che mi interessano ***
            strFields = oCldHhId.FieldsList.Item("ORDER-REFERENCE")
            strFields &= ", " & oCldHhId.FieldsList.Item("ORDER-CREATION-DATE")
            strFields &= ", " & oCldHhId.FieldsList.Item("CUSTOMER-EMAIL")
            strFields &= ", " & oCldHhId.FieldsList.Item("INVOICE-ADDRESS-FIRSTNAME")
            strFields &= ", " & oCldHhId.FieldsList.Item("INVOICE-ADDRESS-LASTNAME")
            strFields &= ", " & oCldHhId.FieldsList.Item("DELIVERY-ADDRESS-PHONE")
            strFields &= ", " & oCldHhId.FieldsList.Item("DELIVERY-ADDRESS-DNI")
            strFields &= ", " & oCldHhId.FieldsList.Item("DELIVERY-ADDRESS1")
            strFields &= ", " & oCldHhId.FieldsList.Item("DELIVERY-ADDRESS2")
            strFields &= ", " & oCldHhId.FieldsList.Item("DELIVERY-ADDRESS-POSTCODE")
            strFields &= ", " & oCldHhId.FieldsList.Item("DELIVERY-ADDRESS-CITY")
            strFields &= ", " & oCldHhId.FieldsList.Item("DELIVERY-ADDRESS-STATE-NAME")
            strFields &= ", " & oCldHhId.FieldsList.Item("DELIVERY-ADDRESS-COUNTRY-NAME")
            strFields &= ", " & oCldHhId.FieldsList.Item("TOTAL-SHIPPING-TAX-EXCLUDED")
            strFields &= ", " & oCldHhId.FieldsList.Item("TOTAL-SHIPPING-TAX-INCLUDED")

            '*** Leggo i dati dal testo ***
            'oCldHhId.GetDataFromCSV(strPath.Replace(strPath.Split("\"c).Last, ""), strPath.Split("\"c).Last, "TESTA", "*", "", "", "", "", dsOut)
            oCldHhId.LetturaExcel(strPath, "Testa", strTipo, strFields, "", strFields, oCldHhId.FieldsList.Item("ORDER-CREATION-DATE"), "", dsOut)

            '*** Compongo la select, per selezionare solo i dati che mi interessano ***
            strFields = oCldHhId.FieldsList.Item("ORDER-REFERENCE")
            strFields &= ", " & oCldHhId.FieldsList.Item("ORDER-CREATION-DATE")
            strFields &= ", " & oCldHhId.FieldsList.Item("CUSTOMER-EMAIL")
            strFields &= ", " & oCldHhId.FieldsList.Item("INVOICE-ADDRESS-FIRSTNAME")
            strFields &= ", " & oCldHhId.FieldsList.Item("INVOICE-ADDRESS-LASTNAME")
            strFields &= ", " & oCldHhId.FieldsList.Item("DELIVERY-ADDRESS-PHONE")
            strFields &= ", " & oCldHhId.FieldsList.Item("DELIVERY-ADDRESS-DNI")
            strFields &= ", " & oCldHhId.FieldsList.Item("DELIVERY-ADDRESS1")
            strFields &= ", " & oCldHhId.FieldsList.Item("DELIVERY-ADDRESS2")
            strFields &= ", " & oCldHhId.FieldsList.Item("DELIVERY-ADDRESS-POSTCODE")
            strFields &= ", " & oCldHhId.FieldsList.Item("DELIVERY-ADDRESS-CITY")
            strFields &= ", " & oCldHhId.FieldsList.Item("DELIVERY-ADDRESS-STATE-NAME")
            strFields &= ", " & oCldHhId.FieldsList.Item("DELIVERY-ADDRESS-COUNTRY-NAME")
            strFields &= ", " & oCldHhId.FieldsList.Item("TOTAL-SHIPPING-TAX-EXCLUDED")
            strFields &= ", " & oCldHhId.FieldsList.Item("TOTAL-SHIPPING-TAX-INCLUDED")

            strFields &= ", " & oCldHhId.FieldsList.Item("COMBINATION-EAN13")
            strFields &= ", " & oCldHhId.FieldsList.Item("RIFERIMENTO-DEL-PRODOTTO")
            strFields &= ", " & oCldHhId.FieldsList.Item("PRODUCT-NAME-WITH-COMBINATION")
            strFields &= ", " & oCldHhId.FieldsList.Item("PRODUCT-QUANTITY")
            strFields &= ", " & oCldHhId.FieldsList.Item("PRODUCT-UNIT-PRICE-TAX-INCLUDED")
            strFields &= ", " & oCldHhId.FieldsList.Item("PRODUCT-UNIT-PRICE-TAX-EXCLUDED")
            strFields &= ", " & oCldHhId.FieldsList.Item("TOTAL-DISCOUNTS-TAX-INCLUDED")
            strFields &= ", " & oCldHhId.FieldsList.Item("TOTAL-DISCOUNTS-TAX-EXCLUDED")

            strFields &= ", " & oCldHhId.FieldsList.Item("INVOICETAXTYPE-TAXID-TAXNAME-TAXAMOUNT")

            '*** Leggo i dati dal testo ***
            'oCldHhId.LetturaTesto(strPath, "Corpo", strFields, Nothing, Nothing, dsOut)
            oCldHhId.LetturaExcel(strPath, "Corpo", strTipo, strFields, "", strFields, "", "", dsOut)

            Return True
        Catch ex As Exception
            '---------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            Return False
            '---------------------------------------------------------
        End Try
    End Function
    Public Overridable Function IsFileOpen(ByVal strPath As String, ByRef strOut As String) As Boolean
        Dim stream As FileStream = Nothing
        Dim file As FileInfo
        Try
            file = New FileInfo(strPath)
            stream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None)
            stream.Close()
            strOut = ""
            Return True
        Catch ex As Exception
            If TypeOf ex Is IOException AndAlso IsFileLocked(ex) Then
                ' do something here, either close the file if you have a handle, show a msgbox, retry  or as a last resort terminate the process - which could cause corruption and lose data
                strOut = ex.Message
                Return False
            End If
        End Try
    End Function
    Public Overridable Function IsFileLocked(exception As Exception) As Boolean
        Dim errorCode As Integer = Marshal.GetHRForException(exception) And ((1 << 16) - 1)
        Return errorCode = 32 OrElse errorCode = 33
    End Function
#End Region

#Region "Utility"
    '*** Conversione
    Public Overridable Function ConvertCSVtoXLSX(csvFilePath As String) As String
        Try
            ' Verifica se il file CSV esiste
            If Not System.IO.File.Exists(csvFilePath) Then
                ScriviLog("Il file CSV specificato non esiste.")
                oApp.MsgBoxInfo("Il file CSV specificato non esiste.")
                Return ""
            End If

            ' Ottieni il percorso della cartella del file CSV
            Dim csvFolder As String = System.IO.Path.GetDirectoryName(csvFilePath)

            ' Ottieni il nome del file senza estensione
            Dim fileNameWithoutExtension As String = System.IO.Path.GetFileNameWithoutExtension(csvFilePath)

            ' Costruisci il percorso del file XLSX nella stessa cartella
            Dim xlsxFilePath As String = System.IO.Path.Combine(csvFolder, fileNameWithoutExtension & ".xlsx")

            ' Creare un'applicazione Excel
            Dim excelApp As New Microsoft.Office.Interop.Excel.Application()

            ' Creare un nuovo libro di lavoro
            Dim workbook As Microsoft.Office.Interop.Excel.Workbook = excelApp.Workbooks.Add()

            ' Ottenere il foglio di lavoro attivo
            Dim worksheet As Microsoft.Office.Interop.Excel.Worksheet = CType(workbook.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)

            ' Assegna il nome "Foglio1" al foglio di lavoro
            worksheet.Name = "Foglio1"

            ' Leggere i dati dal file CSV e scriverli nel foglio di lavoro Excel
            Dim lines As String() = System.IO.File.ReadAllLines(csvFilePath)
            For i As Integer = 0 To lines.Length - 1
                Dim line As String = lines(i)
                Dim values As String() = line.Split(";"c)

                For j As Integer = 0 To values.Length - 1
                    ' Incrementare la colonna di 1 perché Excel usa l'indice basato su 1
                    Dim cellAddress As String = GetExcelColumnName(j + 1) & (i + 1)
                    worksheet.Range(cellAddress).Value = values(j).TrimStart(""""c).TrimEnd(""""c).Trim
                Next
            Next

            ' Salvare il libro di lavoro come file XLSX
            workbook.SaveAs(xlsxFilePath)

            ' Chiudere l'applicazione Excel
            excelApp.Quit()

            ' Rilasciare le risorse
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp)

            Return xlsxFilePath
        Catch ex As Exception
            '---------------------------------------------------------
            ScriviLog(ex.Message)
            CLN__STD.GestErr(ex, Me, "")
            '---------------------------------------------------------
        End Try

    End Function
    Public Overridable Function GetExcelColumnName(columnNumber As Integer) As String
        Dim dividend As Integer = columnNumber
        Dim columnName As String = String.Empty

        While dividend > 0
            Dim modulo As Integer = (dividend - 1) Mod 26
            columnName = Convert.ToChar(65 + modulo).ToString() & columnName
            dividend = CInt((dividend - modulo) / 26)
        End While

        Return columnName
    End Function
    '*** Ricerca ***
    Public Overridable Function GetCodartByCodAlt(ByVal strCodAlt As String) As String
        Dim strSQL As String
        Dim dttTmp As DataTable
        Dim strCodart As String
        Try
            strSQL = $"Select * " &
                $" FROM artico " &
                $" WHERE codditt = {CStrSQL(strDittaCorrente)}" &
                $" And ar_codalt = {CStrSQL(strCodAlt)}"
            dttTmp = oCldHhId.OpenRecordset(strSQL, CLE__APP.DBTIPO.DBAZI)
            If dttTmp.Rows.Count > 0 Then
                strCodart = NTSCStr(dttTmp.Rows(0)!ar_codart)
            Else
                strCodart = ""
            End If

            Return strCodart
        Catch ex As Exception
            '---------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '---------------------------------------------------------
        End Try
    End Function
    Public Overridable Function GetCodMastroClienti(ByRef Optional nConto As Integer = 0) As Integer
        Dim strSQL As String = String.Empty
        Dim dttTmp As DataTable
        Dim nCodMast As Integer = 0
        Try
            '*** Controllo se sforo con il progressivo ***
            'strSQL = "Select * FROM tabnuma	" & vbCrLf &
            '            " WHERE tb_numtipo = 'CC'" & vbCrLf & 'Imposto il tipo su Clienti
            '            " AND tb_numprog < 9999" & vbCrLf & 'Controllo che il progressivo non abbia sforato
            '            " ORDER BY tb_numcodl ASC" 'Ordino in modo crescente così il primo sarà il mastro da usare
            'dttTmp = oCldShop.OpenRecordset(strSQL, CLE__APP.DBTIPO.DBAZI, "TABNUMA").Tables("TABNUMA").Copy


            strSQL = "SELECT tabnuma.*" & vbCrLf &
                        ", ISNULL((anagra.an_conto + 1), tb_numcodl * 10000 + 1) AS xx_conto" & vbCrLf &
                        ", ISNULL((SELECT TOP 1 an_conto FROM anagra anagraNEW WHERE anagraNEW.codditt = anagra.codditt AND anagraNEW.an_conto = ISNULL((anagra.an_conto + 1), tb_numcodl * 10000 + 1)), 0)" & vbCrLf &
                        "FROM tabnuma" & vbCrLf &
                        "LEFT OUTER JOIN anagra ON anagra.codditt = tabnuma.codditt " & vbCrLf &
                        "    AND an_codmast = tb_numcodl" & vbCrLf &
                        "    AND  (anagra.an_conto + 1)/10000 = tb_numcodl" & vbCrLf &
                        $" WHERE tabnuma.codditt = {CStrSQL(strDittaCorrente)}" & vbCrLf &
                        "    AND tb_numtipo = 'CC'" & vbCrLf &
                        "    AND ISNULL((SELECT TOP 1 an_conto FROM anagra anagraNEW WHERE anagraNEW.codditt = anagra.codditt AND anagraNEW.an_conto = ISNULL((anagra.an_conto + 1), tb_numcodl * 10000 + 1)), 0) = 0" & vbCrLf &
                        "ORDER BY tb_numcodl, an_conto"
            dttTmp = oCldHhId.OpenRecordset(strSQL, CLE__APP.DBTIPO.DBAZI, "TABNUMA").Tables("TABNUMA").Copy

            If dttTmp.Rows.Count = 0 Then
                ThrowRemoteEvent(New NTSEventArgs(ThMsg.MSG_ERROR, "Non sono stati trovati altri codici mastro da utilizzare!"))
                Return 0
            ElseIf dttTmp.Select($"tb_numcodl = {nCodMast}").Length = 0 Then
                '*** Ottengo il nuovo codice mastro ***
                nCodMast = NTSCInt(dttTmp.Rows(0)!tb_numcodl)
            End If

            nConto = NTSCInt(dttTmp.Rows(0)!xx_conto)

            Return nCodMast
        Catch ex As Exception
            '---------------------------------------------------------
            ScriviLog("Codice mastro clienti non trovato!")
            CLN__STD.GestErr(ex, Me, "")
            Return 0
            '---------------------------------------------------------
        End Try
    End Function
    '*** Controlli  ***
    Public Overridable Function ControlloEsistenzaFoglioExcel(PathfileExcel As String, sNomefoglio As String) As Boolean
        Dim bReturn As Boolean = False
        Dim excelApp As Microsoft.Office.Interop.Excel.Application
        Dim FileExcel As Microsoft.Office.Interop.Excel.Workbook
        Dim FoglioExcel As Microsoft.Office.Interop.Excel.Worksheet
        Dim RangeExcel As Microsoft.Office.Interop.Excel.Range
        Try
            'applicazione Excel
            excelApp = New Microsoft.Office.Interop.Excel.Application
            'cartella di lavoro Excel
            FileExcel = excelApp.Workbooks.Open(PathfileExcel)
            excelApp.Visible = False
            FileExcel.Activate()
            For Each displayWorksheet As Microsoft.Office.Interop.Excel.Worksheet In FileExcel.Worksheets
                'MsgBox(displayWorksheet.Name)
                If displayWorksheet.Name = sNomefoglio Then
                    bReturn = True
                End If
            Next displayWorksheet
            RangeExcel = Nothing
            FoglioExcel = Nothing
            FileExcel = Nothing
            excelApp.Quit()
            excelApp = Nothing
            Return bReturn
        Catch ex As Exception
            RangeExcel = Nothing
            FoglioExcel = Nothing
            FileExcel = Nothing
            excelApp = Nothing
            If ex.Message.Contains("Impossibile eseguire il cast di oggetti COM di tipo") Then Return True
            Return False
        End Try

    End Function
    Public Overridable Function ControllaValiditaFile(ByVal strPath As String, ByVal strNomeFoglio As String, ByRef Optional strOut As String = "") As Boolean
        Try
            '*** In caso non abbia selezionato un file ***
            If String.IsNullOrEmpty(strPath) OrElse String.IsNullOrEmpty(strPath.Trim) Then
                strOut = "Il file selezionato non trovato!"
                Return False
            End If

            '*** in caso il file selezionato non esista ***
            If Not My.Computer.FileSystem.FileExists(strPath) Then
                strOut = "Il file selezionato non trovato!"
                Return False
            End If

            '*** in caso il file selezionato non esista ***
            If Not IsFileOpen(strPath, strOut) Then Return False

            '*** Controllo se esiste il foglio desiderato ***
            If Not ControlloEsistenzaFoglioExcel(strPath, strNomeFoglio) Then
                strOut = "Il nome del foglio nel file selezionato non è compatibile per l'importazione scelta!"
                Return False
            End If

            Return True
        Catch ex As Exception
            '---------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '---------------------------------------------------------
        End Try
    End Function
    '*** Gestione log ***
    Public Overridable Sub ScriviLog(ByVal strText As String)
        Try
            oCldHhId.ScriviLog(strText)
        Catch ex As Exception
            '---------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '---------------------------------------------------------
        End Try
    End Sub

    '*** Dati Indirizzo ***
    Public Overridable Function GetComuneByDescr(ByVal strComuni As String) As String
        Dim StrSQL As String
        Dim dttIndir As DataTable
        Try
            If Not String.IsNullOrEmpty(strComuni) Then
                StrSQL = "SELECT * FROM comuni WHERE co_denom = " & CStrSQL(strComuni).ToUpper
                dttIndir = ocldBase.OpenRecordset(StrSQL, CLE__APP.DBTIPO.DBAZI)
                If dttIndir.Rows.Count = 0 Then Return ""
                Return NTSCStr(dttIndir.Rows(0)!co_codcomu)
            End If

            Return ""
        Catch ex As Exception
            '---------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            Return ""
            '---------------------------------------------------------
        End Try
    End Function
    Public Overridable Function GetStatoByDescr(ByVal strStato As String) As String
        Dim StrSQL As String
        Dim dttIndir As DataTable
        Try
            If Not String.IsNullOrEmpty(strStato) Then
                StrSQL = "SELECT * FROM tabprov WHERE tb_desprov = " & CStrSQL(strStato).ToUpper & " OR tb_codprov = " & CStrSQL(strStato).ToUpper
                dttIndir = ocldBase.OpenRecordset(StrSQL, CLE__APP.DBTIPO.DBAZI)
                If dttIndir.Rows.Count = 0 Then Return ""
                Return NTSCStr(dttIndir.Rows(0)!tb_codprov)
            End If

            Return ""
        Catch ex As Exception
            '---------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            Return ""
            '---------------------------------------------------------
        End Try
    End Function
    Public Overridable Function GetNazioneByDescr(ByVal strNazione As String) As String
        Dim StrSQL As String
        Dim dttIndir As DataTable
        Try
            If Not String.IsNullOrEmpty(strNazione) Then
                StrSQL = "SELECT * FROM tabstat WHERE tb_desstat = " & CStrSQL(strNazione).ToUpper & " OR " & " tb_siglacee = " & CStrSQL(strNazione).ToUpper
                dttIndir = ocldBase.OpenRecordset(StrSQL, CLE__APP.DBTIPO.DBAZI)
                If dttIndir.Rows.Count = 0 Then Return ""
                Return NTSCStr(dttIndir.Rows(0)!tb_codstat)
            End If

            Return ""
        Catch ex As Exception
            '---------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            Return ""
            '---------------------------------------------------------
        End Try
    End Function
    Public Overridable Function GetCoddestByDescr(ByVal oImpegno As DTO_HHID_TestataImp, ByVal lConto As Integer) As Integer
        Dim strIndir As String = String.Empty
        Dim strComune As String = String.Empty
        Dim strStato As String = String.Empty
        Dim strNazione As String = String.Empty

        Dim StrWhere As String = String.Empty
        Dim dttIndir As DataTable
        Try
            '*** Ottengo l'indirizzo ***
            strIndir = NTSCStr(oImpegno.oCliente.strIndir)
            strComune = NTSCStr(oImpegno.oCliente.strCodComune)
            strStato = NTSCStr(oImpegno.oCliente.strCodStato)
            strNazione = NTSCStr(oImpegno.oCliente.strCodNazione)

            StrWhere = $"dd_conto = {CStrSQL(lConto)}" & vbCrLf
            StrWhere &= $" AND dd_inddest = {CStrSQL(strIndir)}" & vbCrLf
            If Not String.IsNullOrEmpty(strComune) Then StrWhere &= $" AND dd_codcomu = {CStrSQL(strComune)}" & vbCrLf
            If Not String.IsNullOrEmpty(strStato) Then StrWhere &= $" AND dd_prodest = {CStrSQL(strStato)}" & vbCrLf
            If Not String.IsNullOrEmpty(strNazione) Then StrWhere &= $" AND dd_stato = {CStrSQL(strNazione)}" & vbCrLf

            dttIndir = oCldHhId.OpenRecordset($"SELECT * FROM destdiv WHERE codditt = {CStrSQL(strDittaCorrente)} AND " & StrWhere, CLE__APP.DBTIPO.DBAZI)
            If dttIndir.Rows.Count = 0 Then Return 0

            Return NTSCInt(dttIndir.Rows(0)!dd_coddest)
        Catch ex As Exception
            '---------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '---------------------------------------------------------
        End Try
    End Function
#End Region


#Region "Datatable To DTO"
    Public Overridable Function DataTableAmazonToDTO(ByVal dsHhid As DataSet, ByVal oImpegni As List(Of DTO_HHID_TestataImp)) As List(Of DTO_HHID_TestataImp)
        Dim oImpegno As DTO_HHID_TestataImp
        Dim oCorpo As DTO_HHID_CorpoImp
        Dim oCliente As DTO_HHID_ClienteImp
        Try
            For Each dtrEt As DataRow In dsHhid.Tables("TESTA").Rows
                oImpegno = New DTO_HHID_TestataImp
                oImpegno.strFonte = "Amazon"

                '--- Dati del cliente -----------------------
                oCliente = New DTO_HHID_ClienteImp

                '*** Dati anagrafici ***
                oCliente.strNome = NTSCStr(dtrEt.Item("buyer-name"))
                oCliente.strEmail = NTSCStr(dtrEt.Item("buyer-email"))
                oCliente.strPhone = NTSCStr(dtrEt.Item("buyer-phone-number"))

                '*** Indirizzo ***
                oCliente.strIndir = NTSCStr(dtrEt.Item("ship-address-1")) & " " & NTSCStr(dtrEt.Item("ship-address-2"))
                oCliente.strCitta = NTSCStr(dtrEt.Item("ship-city"))
                oCliente.strCap = NTSCStr(dtrEt.Item("ship-postal-code"))
                oCliente.strCodComune = GetComuneByDescr(NTSCStr(dtrEt.Item("ship-city")))
                oCliente.strCodProv = GetStatoByDescr(NTSCStr(dtrEt.Item("ship-state")))
                oCliente.strCodStato = GetStatoByDescr(NTSCStr(dtrEt.Item("ship-state")))
                oCliente.strCodNazione = GetNazioneByDescr(NTSCStr(dtrEt.Item("ship-country")))

                oImpegno.oCliente = oCliente
                '--- Dati della testata ---------------------
                oImpegno.strTipoOrd = "R"
                oImpegno.nAnnoOrd = NTSCDate(dtrEt.Item("purchase-date")).Year
                oImpegno.strSerieOrd = strSerie_Impegni_Amazon
                oImpegno.lNumTmpOrd = 0

                oImpegno.dDatDoc = NTSCDate(dtrEt.Item("purchase-date"))
                oImpegno.strRiferim = NTSCStr(dtrEt.Item("order-id")).Trim

                oImpegno.nScoPag = 0
                oImpegno.strScorporo = NTSCStr(IIf(bPrezzoIvato, "S", "N"))

                oImpegno.nTipoBf = nTipoBF_Impegni
                oImpegno.nCodMaga = nMagaz_Impegni
                oImpegno.nListino = nListino_Impegni
                oImpegno.nCodPaga = nPaga_Impegni

                oImpegno.nSpeAcc = NTSCDec(NTSCStr(dtrEt.Item("shipping-price")).Replace(".", ",")) - NTSCDec(NTSCStr(dtrEt.Item("shipping-tax")).Replace(".", ","))

                '--- Dati della Corpo -----------------------
                For Each dtrEc As DataRow In dsHhid.Tables("CORPO").Select($"[order-id] = {CStrSQL(dtrEt.Item("order-id"))}")
                    Try
                        oCorpo = New DTO_HHID_CorpoImp

                        oCorpo.nRiga = 0
                        oCorpo.strCodAlt = NTSCStr(dtrEc.Item("sku"))
                        oCorpo.strCodArt = GetCodartByCodAlt(oCorpo.strCodAlt)
                        oCorpo.strDescr = NTSCStr(dtrEc.Item("product-name"))

                        oCorpo.dQuant = NTSCDec(NTSCStr(dtrEc.Item("quantity-purchased")).Replace(".", ","))

                        If bPrezzoIvato Then
                            oCorpo.dPrezzoIva = NTSCDec(NTSCStr(dtrEc.Item("item-price")).Replace(".", ",")) / NTSCDec(NTSCStr(dtrEc.Item("quantity-purchased")).Replace(".", ","))
                        Else
                            oCorpo.dPrezzo = NTSCDec(NTSCStr(dtrEc.Item("item-price")).Replace(".", ",")) / NTSCDec(NTSCStr(dtrEc.Item("quantity-purchased")).Replace(".", ","))
                        End If
                        oCorpo.nControp = nCoVend_Impegni
                        oCorpo.nCodIva = nCodIva_Impegni

                        oCorpo.dPrzPromo = NTSCDec(NTSCStr(dtrEc.Item("item-promotion-discount")).Replace(".", ","))
                        oCorpo.nContropPromo = nPromoCoVend_Impegni
                        oCorpo.nCodIvaPromo = nPromoCodIva_Impegni
                        oCorpo.strTipoPromo = "ART"

                        oImpegno.oCorpo.Add(oCorpo)
                    Catch ex As Exception
                        ScriviLog("/!\ Riga non conforme!")
                        ScriviLog(ex.Message)
                    End Try
                Next

                oImpegni.Add(oImpegno)
            Next

            Return oImpegni
        Catch ex As Exception
            '---------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '---------------------------------------------------------
        End Try
    End Function
    Public Overridable Function DataTableXbioToDTO(ByVal dsHhid As DataSet, ByVal oImpegni As List(Of DTO_HHID_TestataImp)) As List(Of DTO_HHID_TestataImp)
        Dim oImpegno As DTO_HHID_TestataImp
        Dim oCorpo As DTO_HHID_CorpoImp
        Dim oCliente As DTO_HHID_ClienteImp
        Try
            For Each dtrEt As DataRow In dsHhid.Tables("TESTA").Rows
                oImpegno = New DTO_HHID_TestataImp
                oImpegno.strFonte = "Xbio"

                '--- Dati del cliente -----------------------
                oCliente = New DTO_HHID_ClienteImp

                '*** Dati anagrafici ***
                oCliente.strNome = NTSCStr(NTSCStr(dtrEt.Item("Invoice-Address-Firstname")) & " " & NTSCStr(dtrEt.Item("Invoice-Address-Lastname")))
                oCliente.strEmail = NTSCStr(dtrEt.Item("Customer-Email"))
                oCliente.strPhone = NTSCStr(dtrEt.Item("Delivery-Address-Phone"))

                '*** Indirizzo ***
                oCliente.strIndir = NTSCStr(dtrEt.Item("delivery-address1")) & " " & NTSCStr(dtrEt.Item("delivery-address2"))
                oCliente.strCitta = NTSCStr(dtrEt.Item("DELIVERY-ADDRESS-CITY"))
                oCliente.strCap = NTSCStr(dtrEt.Item("DELIVERY-ADDRESS-POSTCODE"))
                oCliente.strCodComune = GetComuneByDescr(NTSCStr(dtrEt.Item("DELIVERY-ADDRESS-CITY")))
                oCliente.strCodProv = GetStatoByDescr(NTSCStr(dtrEt.Item("DELIVERY-ADDRESS-STATE-NAME")))
                oCliente.strCodStato = GetStatoByDescr(NTSCStr(dtrEt.Item("DELIVERY-ADDRESS-STATE-NAME")))
                oCliente.strCodNazione = GetNazioneByDescr(NTSCStr(dtrEt.Item("DELIVERY-ADDRESS-COUNTRY-NAME")))

                oImpegno.oCliente = oCliente
                '--- Dati della testata ---------------------
                oImpegno.strTipoOrd = "R"
                oImpegno.nAnnoOrd = NTSCDate(dtrEt.Item("ORDER-CREATION-DATE")).Year
                oImpegno.strSerieOrd = strSerie_Impegni_Xbio
                oImpegno.lNumTmpOrd = 0

                oImpegno.dDatDoc = NTSCDate(dtrEt.Item("ORDER-CREATION-DATE"))
                oImpegno.strRiferim = NTSCStr(dtrEt.Item("ORDER-REFERENCE")).Trim

                oImpegno.nScoPag = 0
                oImpegno.strScorporo = NTSCStr(IIf(bPrezzoIvato, "S", "N"))

                oImpegno.nTipoBf = nTipoBF_Impegni
                oImpegno.nCodMaga = nMagaz_Impegni
                oImpegno.nListino = nListino_Impegni
                oImpegno.nCodPaga = nPaga_Impegni

                oImpegno.nSpeAcc = NTSCDec(NTSCStr(dtrEt.Item("TOTAL-SHIPPING-TAX-EXCLUDED")).Replace(".", ","))

                '--- Dati della Corpo -----------------------
                For Each dtrEc As DataRow In dsHhid.Tables("CORPO").Select($"[ORDER-REFERENCE] = {CStrSQL(dtrEt.Item("ORDER-REFERENCE"))}")
                    oCorpo = New DTO_HHID_CorpoImp

                    oCorpo.nRiga = 0
                    oCorpo.strCodAlt = ""
                    oCorpo.strCodArt = NTSCStr(dtrEc.Item("RIFERIMENTO-DEL-PRODOTTO"))
                    oCorpo.strDescr = NTSCStr(dtrEc.Item("PRODUCT-NAME-WITH-COMBINATION"))

                    oCorpo.dQuant = NTSCDec(NTSCStr(dtrEc.Item("PRODUCT-QUANTITY")).Replace(".", ","))

                    oCorpo.dPrezzoIva = NTSCDec(NTSCStr(dtrEc.Item("PRODUCT-UNIT-PRICE-TAX-INCLUDED")).Replace(".", ",")) / NTSCDec(NTSCStr(dtrEc.Item("PRODUCT-QUANTITY")).Replace(".", ","))
                    oCorpo.dPrezzo = NTSCDec(NTSCStr(dtrEc.Item("PRODUCT-UNIT-PRICE-TAX-EXCLUDED")).Replace(".", ",")) / NTSCDec(NTSCStr(dtrEc.Item("PRODUCT-QUANTITY")).Replace(".", ","))
                    oCorpo.nControp = nCoVend_Impegni
                    oCorpo.nCodIva = nCodIva_Impegni

                    '--- Dati del Corpo -------------------------
                    Dim dtrHhid As List(Of DataRow) = dsHhid.Tables("CORPO").Select($"[ORDER-REFERENCE] = {CStrSQL(dtrEt.Item("ORDER-REFERENCE"))}").ToList
                    If dtrHhid.IndexOf(dtrEc) = dtrHhid.Count - 1 Then
                        If bPrezzoIvato Then
                            oCorpo.dPrzPromo = NTSCDec(NTSCStr(dtrEc.Item("Total-Discounts-Tax-included")).Replace(".", ","))
                        Else
                            oCorpo.dPrzPromo = NTSCDec(NTSCStr(dtrEc.Item("Total-Discounts-Tax-excluded")).Replace(".", ","))
                        End If
                        oCorpo.nContropPromo = nPromoCoVend_Impegni
                        oCorpo.nCodIvaPromo = nPromoCodIva_Impegni
                        oCorpo.strTipoPromo = "TOT"
                    End If

                    oImpegno.oCorpo.Add(oCorpo)
                Next

                oImpegni.Add(oImpegno)
            Next

            Return oImpegni
        Catch ex As Exception
            '---------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '---------------------------------------------------------
        End Try
    End Function
    Public Overridable Function DataTableSleepHToDTO(ByVal dsHhid As DataSet, ByVal oImpegni As List(Of DTO_HHID_TestataImp)) As List(Of DTO_HHID_TestataImp)
        Dim oImpegno As DTO_HHID_TestataImp
        Dim oCorpo As DTO_HHID_CorpoImp
        Dim oCliente As DTO_HHID_ClienteImp
        Try
            For Each dtrEt As DataRow In dsHhid.Tables("TESTA").Rows
                oImpegno = New DTO_HHID_TestataImp
                oImpegno.strFonte = "SleepH"

                '--- Dati del cliente -----------------------
                oCliente = New DTO_HHID_ClienteImp

                '*** Dati anagrafici ***
                oCliente.strNome = NTSCStr(NTSCStr(dtrEt.Item("Invoice-Address-Firstname")) & " " & NTSCStr(dtrEt.Item("Invoice-Address-Lastname")))
                oCliente.strEmail = NTSCStr(dtrEt.Item("Customer-Email"))
                oCliente.strPhone = NTSCStr(dtrEt.Item("Delivery-Address-Phone"))

                '*** Indirizzo ***
                oCliente.strIndir = NTSCStr(dtrEt.Item("delivery-address1")) & " " & NTSCStr(dtrEt.Item("delivery-address2"))
                oCliente.strCitta = NTSCStr(dtrEt.Item("DELIVERY-ADDRESS-CITY"))
                oCliente.strCap = NTSCStr(dtrEt.Item("DELIVERY-ADDRESS-POSTCODE"))
                oCliente.strCodComune = GetComuneByDescr(NTSCStr(dtrEt.Item("DELIVERY-ADDRESS-CITY")))
                oCliente.strCodProv = GetStatoByDescr(NTSCStr(dtrEt.Item("DELIVERY-ADDRESS-STATE-NAME")))
                oCliente.strCodStato = GetStatoByDescr(NTSCStr(dtrEt.Item("DELIVERY-ADDRESS-STATE-NAME")))
                oCliente.strCodNazione = GetNazioneByDescr(NTSCStr(dtrEt.Item("DELIVERY-ADDRESS-COUNTRY-NAME")))

                oImpegno.oCliente = oCliente
                '--- Dati della testata ---------------------
                oImpegno.strTipoOrd = "R"
                oImpegno.nAnnoOrd = NTSCDate(dtrEt.Item("ORDER-CREATION-DATE")).Year
                oImpegno.strSerieOrd = strSerie_Impegni_SleepH
                oImpegno.lNumTmpOrd = 0

                oImpegno.dDatDoc = NTSCDate(dtrEt.Item("ORDER-CREATION-DATE"))
                oImpegno.strRiferim = NTSCStr(dtrEt.Item("ORDER-REFERENCE")).Trim

                oImpegno.nScoPag = 0
                oImpegno.strScorporo = NTSCStr(IIf(bPrezzoIvato, "S", "N"))

                oImpegno.nTipoBf = nTipoBF_Impegni
                oImpegno.nCodMaga = nMagaz_Impegni
                oImpegno.nListino = nListino_Impegni
                oImpegno.nCodPaga = nPaga_Impegni

                oImpegno.nSpeAcc = NTSCDec(NTSCStr(dtrEt.Item("TOTAL-SHIPPING-TAX-EXCLUDED")).Replace(".", ","))

                '--- Dati del Corpo -------------------------
                Dim dtrHhid As List(Of DataRow) = dsHhid.Tables("CORPO").Select($"[ORDER-REFERENCE] = {CStrSQL(dtrEt.Item("ORDER-REFERENCE"))}").ToList
                For Each dtrEc As DataRow In dtrHhid
                    oCorpo = New DTO_HHID_CorpoImp

                    oCorpo.nRiga = 0
                    oCorpo.strCodAlt = ""
                    oCorpo.strCodArt = NTSCStr(dtrEc.Item("RIFERIMENTO-DEL-PRODOTTO"))
                    oCorpo.strDescr = NTSCStr(dtrEc.Item("PRODUCT-NAME-WITH-COMBINATION"))

                    oCorpo.dQuant = NTSCDec(NTSCStr(dtrEc.Item("PRODUCT-QUANTITY")).Replace(".", ","))

                    oCorpo.dPrezzoIva = NTSCDec(NTSCStr(dtrEc.Item("PRODUCT-UNIT-PRICE-TAX-INCLUDED")).Replace(".", ",")) / NTSCDec(NTSCStr(dtrEc.Item("PRODUCT-QUANTITY")).Replace(".", ","))
                    oCorpo.dPrezzo = NTSCDec(NTSCStr(dtrEc.Item("PRODUCT-UNIT-PRICE-TAX-EXCLUDED")).Replace(".", ",")) / NTSCDec(NTSCStr(dtrEc.Item("PRODUCT-QUANTITY")).Replace(".", ","))
                    oCorpo.nControp = nCoVend_Impegni
                    oCorpo.nCodIva = nCodIva_Impegni

                    If dtrHhid.IndexOf(dtrEc) = dtrHhid.Count - 1 Then
                        If bPrezzoIvato Then
                            oCorpo.dPrzPromo = NTSCDec(NTSCStr(dtrEc.Item("Total-Discounts-Tax-included")).Replace(".", ","))
                        Else
                            oCorpo.dPrzPromo = NTSCDec(NTSCStr(dtrEc.Item("Total-Discounts-Tax-excluded")).Replace(".", ","))
                        End If
                        oCorpo.nContropPromo = nPromoCoVend_Impegni
                        oCorpo.nCodIvaPromo = nPromoCodIva_Impegni
                        oCorpo.strTipoPromo = "TOT"
                    End If

                    oImpegno.oCorpo.Add(oCorpo)
                Next

                oImpegni.Add(oImpegno)
            Next

            Return oImpegni
        Catch ex As Exception
            '---------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '---------------------------------------------------------
        End Try
    End Function
#End Region
End Class

