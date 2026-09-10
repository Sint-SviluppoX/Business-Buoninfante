Imports System.Data
Imports NTSInformatica.CLN__STD
Imports System.Globalization
Imports System
Public Class CLFMGETTE
    Inherits CLEMGETTE
    Public _oClhEtte As CLHMGETTE
    Public Property oClhEtte() As CLHMGETTE
        Get
            If _oClhEtte Is Nothing Then _oClhEtte = CType(oCldEtte, CLHMGETTE)
            Return _oClhEtte
        End Get
        Set(ByVal value As CLHMGETTE)
            _oClhEtte = value
        End Set
    End Property

#Region "RFid - Aly"

    Public Overridable Function InserisciEtichetteDBExt(ByVal strTipoDoc As String, ByVal nAnnoDoc As Integer, ByVal strSerieDoc As String, ByVal lNumDoc As Integer, ByVal nRigaDoc As Integer, ByVal DataDoc As Date, ByVal CodArt As String, ByVal DesArt As String, ByVal Qta As Double, ByVal nEticInCorso As Integer) As Boolean
        Try
            ThrowRemoteEvent(New NTSEventArgs("LABEL", oApp.Tr(Me, 129030099509236239, "Preparazione etichetta RFID n° |" & nEticInCorso & "| in corso...")))
            Return oClhEtte.InserisciEtichetteDBExt(strTipoDoc, nAnnoDoc, strSerieDoc, lNumDoc, nRigaDoc, DataDoc, CodArt, DesArt, Qta)
        Catch ex As Exception
            '-------------------------------------------------
            Dim strErr As String = CLN__STD.GestError(ex, Me, "", oApp.InfoError, oApp.ErrorLogFile, True)
            '-------------------------------------------------
        End Try
    End Function
    Public Overridable Function CancellaEtichetteDBExt(ByVal strTipoDoc As String, ByVal nAnnoDoc As Integer, ByVal strSerieDoc As String, ByVal lNumDoc As Integer, ByVal nRigaDoc As Integer) As Boolean
        Try
            Return oClhEtte.CancellaEtichetteDBExt(strTipoDoc, nAnnoDoc, strSerieDoc, lNumDoc, nRigaDoc)
        Catch ex As Exception
            '-------------------------------------------------
            Dim strErr As String = CLN__STD.GestError(ex, Me, "", oApp.InfoError, oApp.ErrorLogFile, True)
            '-------------------------------------------------
        End Try
    End Function
    Public Overridable Function CancellaEtichettePrecedentiADataDBExt(ByVal DataRifDelete As Date, ByVal QualeDataRif As String, ByRef NRecordDelete As Integer) As Boolean
        Try
            Return oClhEtte.CancellaEtichettePrecedentiADataDBExt(DataRifDelete, QualeDataRif, NRecordDelete)
        Catch ex As Exception
            '-------------------------------------------------
            Dim strErr As String = CLN__STD.GestError(ex, Me, "", oApp.InfoError, oApp.ErrorLogFile, True)
            '-------------------------------------------------
        End Try
    End Function
    Public Overridable Function MyValCodiceDB(ByVal Codice As String, ByVal strNomeTabella As String, ByVal strTipoCod As String) As String
        Try
            Dim TmpDes As String = String.Empty
            oClhEtte.ValCodiceDb(NTSCStr(Codice).ToString, strDittaCorrente, strNomeTabella, strTipoCod, TmpDes)
            Return TmpDes
        Catch ex As Exception
            '-------------------------------------------------
            Dim strErr As String = CLN__STD.GestError(ex, Me, "", oApp.InfoError, oApp.ErrorLogFile, True)
            Return String.Empty
            '-------------------------------------------------
        End Try
    End Function
    Public Overridable Function SelezionaDaDocMagaRFID(ByVal strDitta As String, ByVal strTipodoc As String, ByVal lDaclie As Integer,
                                                    ByVal lAclie As Integer, ByVal dtDadata As Date, ByVal dtAdata As Date,
                                                    ByVal lDaanno As Integer, ByVal lAanno As Integer, ByVal strDaserie As String,
                                                    ByVal strAserie As String, ByVal lDanumero As Integer, ByVal lAnumero As Integer,
                                                    ByVal lDariga As Integer, ByVal lAriga As Integer, ByVal bDocMagazz As Boolean,
                                                    ByVal bColli As Boolean, ByRef dttOut As DataTable) As Boolean
        Try
            Return oClhEtte.SelezionaDaDocMagaRFID(strDitta, strTipodoc, lDaclie, lAclie, dtDadata, dtAdata,
                                                         lDaanno, lAanno, strDaserie,
                                                         strAserie, lDanumero, lAnumero,
                                                         lDariga, lAriga, bDocMagazz,
                                                         bColli, dttOut)
        Catch ex As Exception
            '-------------------------------------------------
            Dim strErr As String = CLN__STD.GestError(ex, Me, "", oApp.InfoError, oApp.ErrorLogFile, True)
            '-------------------------------------------------
        End Try
    End Function
    Public Overridable Function CancellaEtichetteAllDBExt(ByRef NRecordDelete As Integer) As Boolean
        Try
            Return oClhEtte.CancellaEtichetteAllDBExt(NRecordDelete)
        Catch ex As Exception
            '-------------------------------------------------
            Dim strErr As String = CLN__STD.GestError(ex, Me, "", oApp.InfoError, oApp.ErrorLogFile, True)
            '-------------------------------------------------
        End Try
    End Function
#End Region

    Public Overrides Function RiempiTmpTable(bPrezzi As Boolean, bDatamod As Boolean, dtDatamod As Date, bSolobar As Boolean, strSolobar As String, strUmBC As String, bArtbar As Boolean, strOrigine As String, bBD_DaCodarfo As Boolean, lLista As Integer, bColli As Boolean, bNEtichetteBarCode As Boolean, bPrezzival As Boolean, lListino As Integer, dtValidita As Date, lValuta As Integer, bMoltdiv As Boolean, strQualeOpt As String, bMisura1 As Boolean, lNumMolt As Integer, bDividiQuantConfez As Boolean, lIIstMatr As Integer, lIIstMats As Integer, bUsaTtstMatr As Boolean, lNumEtiche As Integer, nBD_Listino As Integer, bModTCO As Boolean, strTaglia As String, lIIstMatrTmp As Integer, bStampaNegativi As Boolean, ByRef dttOut As DataTable, ByRef dttTmp As DataTable, ByRef dttTmp2 As DataTable, ByRef bDocMagazz As Boolean, ByRef bDaDoc As Boolean, ByRef strCode As String, ByRef strCodartList As String, ByRef strCodart As String, ByRef bAvvisare As Boolean, nEtichetteAggPerTC As Integer, bUsaNumetiDaListaSel As Boolean, bPrzNetClasseSconto As Boolean, bUsaNumetiDaListaSel_DaNote As Boolean) As Boolean
        Dim nProg As Integer = 0
        Dim nWeek As Integer = 0
        Dim strUpdate As String = String.Empty
        Dim strSQL As String = String.Empty
        Dim dttMatr As DataTable
        Dim nRis As Integer
        Try
            If Not MyBase.RiempiTmpTable(bPrezzi, bDatamod, dtDatamod, bSolobar, strSolobar, strUmBC, bArtbar, strOrigine,
                                         bBD_DaCodarfo, lLista, bColli, bNEtichetteBarCode, bPrezzival, lListino, dtValidita,
                                         lValuta, bMoltdiv, strQualeOpt, bMisura1, lNumMolt, bDividiQuantConfez, lIIstMatr,
                                         lIIstMats, bUsaTtstMatr, lNumEtiche, nBD_Listino, bModTCO, strTaglia, lIIstMatrTmp,
                                         bStampaNegativi, dttOut, dttTmp, dttTmp2, bDocMagazz, bDaDoc, strCode, strCodartList,
                                         strCodart, bAvvisare, nEtichetteAggPerTC, bUsaNumetiDaListaSel, bPrzNetClasseSconto, bUsaNumetiDaListaSel_DaNote) Then Return False
            '*** Controllo se è la stampa che mi interessa ***
            If bMoltdiv Then Return True

            strSQL = $"SELECT * FROM TTSTMATR WHERE instid = {CStrSQL(lIIstMatr)}"
            dttMatr = oCldEtte.OpenRecordset(strSQL, CLE__APP.DBTIPO.DBAZI, "TTSTMATR").Tables("TTSTMATR")

            '*** Scorro le righe ***
            For Each dtrMatr As DataRow In dttMatr.Rows
                nWeek = DatePart(DateInterval.WeekOfYear, NTSCDate(dtrMatr!tt_aammgg))
                nProg = NTSCInt(ocldBase.GetSettingBusDitt(strDittaCorrente, "BSMGETTE", "RECENT", ".", "Settimana_" & NTSCDate(dtrMatr!tt_aammgg).Year & "_" & nWeek, "", "", "0")) + 1


                strUpdate = "UPDATE TTSTMATR " &
                    $" SET tt_hhprogr = {CStrSQL(nProg)}, " &
                    $" tt_hhweek = {CStrSQL(nWeek)} " & vbCrLf &
                    $" WHERE instid = {CStrSQL(lIIstMatr)}" & vbCrLf &
                    $" AND tt_codart = {CStrSQL(dtrMatr!tt_codart)}" & vbCrLf &
                    $" AND tt_magaz = {CStrSQL(dtrMatr!tt_magaz)}" & vbCrLf &
                    $" AND tt_aammgg = {CDataSQL(dtrMatr!tt_aammgg)}" & vbCrLf &
                    $" AND tt_tipork = {CStrSQL(dtrMatr!tt_tipork)}" & vbCrLf &
                    $" AND tt_anno = {CStrSQL(dtrMatr!tt_anno)}" & vbCrLf &
                    $" AND tt_serie = {CStrSQL(dtrMatr!tt_serie)}" & vbCrLf &
                    $" AND tt_numdoc = {CStrSQL(dtrMatr!tt_numdoc)}" & vbCrLf &
                    $" AND tt_riga = {CStrSQL(dtrMatr!tt_riga)}" & vbCrLf
                nRis = oCldEtte.Execute(strUpdate, CLE__APP.DBTIPO.DBAZI)
                If nRis > 0 Then ocldBase.SaveSettingBusDitt(strDittaCorrente, "BSMGETTE", "RECENT", ".", "Settimana_" & NTSCDate(dtrMatr!tt_aammgg).Year & "_" & nWeek, nProg.ToString, "", False, True, False)
            Next

            Return True
        Catch ex As Exception
            '---------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '---------------------------------------------------------
        End Try
    End Function

#Region "Generazione QR etichette"

    Public Overridable Function CreaQR(ByVal dttQR As DataTable,
                                       Optional ByVal strPath As String = "") As String
        Try
            If dttQR Is Nothing OrElse dttQR.Rows.Count = 0 Then Return String.Empty

            If String.IsNullOrWhiteSpace(strPath) Then
                strPath = System.IO.Path.Combine(oApp.ServerDir, "Images", "QR")
            End If

            'Una cartella distinta evita interferenze tra stampe contemporanee.
            Dim strCartellaSessione As String = System.IO.Path.Combine(
                strPath,
                DateTime.Now.ToString("yyyyMMdd_HHmmss_fff") & "_" & Guid.NewGuid().ToString("N"))
            System.IO.Directory.CreateDirectory(strCartellaSessione)

            Dim oEncoder As New MessagingToolkit.QRCode.Codec.QRCodeEncoder
            For Each dtrQR As DataRow In dttQR.Rows
                Using img As Bitmap = New Bitmap(oEncoder.Encode(NTSCStr(dtrQR!xx_qr)))
                    img.Save(System.IO.Path.Combine(strCartellaSessione,
                                                    NTSCStr(dtrQR!xx_FileName)),
                             Imaging.ImageFormat.Png)
                End Using
            Next

            Return strCartellaSessione & System.IO.Path.DirectorySeparatorChar
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
            Return String.Empty
        End Try
    End Function

    Public Overridable Function OttieniDatiPerQR(ByVal lInstId As Integer) As DataTable
        Try
            Dim dttQR As DataTable = oClhEtte.OttieniDatiPerQR(strDittaCorrente, lInstId)
            If dttQR Is Nothing Then Return Nothing

            dttQR.Columns.Add("xx_qr", GetType(String))
            dttQR.Columns.Add("xx_FileName", GetType(String))

            For Each dtrQR As DataRow In dttQR.Rows
                dtrQR!xx_qr = "|JF" & NTSCStr(dtrQR!tt_hhweek) &
                    NTSCStr(dtrQR!tt_anno).Substring(2, 2) &
                    NTSCStr(dtrQR!tt_hhprogr).PadLeft(4, "0"c) &
                    "|" & NTSCStr(dtrQR!xx_codarfo).Trim() &
                    "|" & NTSCStr(dtrQR!xx_code).Trim() &
                    "|" & NTSCStr(dtrQR!xx_hhulrif).Trim() &
                    "|" & Right("00" & NTSCStr(dtrQR!tt_hhweek), 2) &
                    NTSCStr(dtrQR!tt_anno) & "|"

                dtrQR!xx_FileName = NTSCStr(dtrQR!tt_tipork) & "_" &
                    NTSCStr(dtrQR!tt_anno) & "_" & NTSCStr(dtrQR!tt_serie).Trim() & "_" &
                    NTSCStr(dtrQR!tt_numdoc) & "_" & NTSCStr(dtrQR!tt_riga) & ".png"
            Next

            Return dttQR
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
            Return Nothing
        End Try
    End Function

#End Region

End Class
