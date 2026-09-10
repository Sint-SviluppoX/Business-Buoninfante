Imports System.Data
Imports System.Data.Common
Imports NTSInformatica.CLN__STD
Imports System
Imports System.Data.SqlClient

Public Class CLHMGETTE
    Inherits CLDMGETTE
    Dim SqlConn As SqlConnection = Nothing
    Dim SqlComm As SqlCommand = Nothing
    Dim SqlDA As SqlDataAdapter = Nothing

#Region "RFid - Alfy"
    Public Overridable Function GetStringaConnessioneDBExt() As String
        GetStringaConnessioneDBExt = String.Empty
        Dim strSQL As String = ""
        Dim dttOut As New DataTable
        Try
            strSQL = "SELECT tb_hhcssqldbext FROM tabanaz WHERE codditt=" & CStrSQL(oApp.Ditta)
            dttOut = OpenRecordset(strSQL, CLE__APP.DBTIPO.DBAZI)
            If dttOut.Rows.Count > 0 AndAlso Not String.IsNullOrEmpty(dttOut.Rows(0)(0).ToString) Then GetStringaConnessioneDBExt = dttOut.Rows(0)(0).ToString.Trim
            If Not TestConnessioneExt(GetStringaConnessioneDBExt) Then Return String.Empty
        Catch ex As Exception
            '--------------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '--------------------------------------------------------------
        End Try
    End Function
    Public Overridable Function TestConnessioneExt(StrConn As String) As Boolean
        Try
            If Not String.IsNullOrEmpty(StrConn) Then
                SqlConn = New SqlConnection(StrConn)
                SqlConn.Open()
                SqlConn.Close()
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
            Return False
        End Try
    End Function
    Public Overridable Function InserisciEtichetteDBExt(ByVal strTipoDoc As String, ByVal nAnnoDoc As Integer, ByVal strSerieDoc As String, ByVal lNumDoc As Integer, ByVal nRigaDoc As Integer, ByVal DataDoc As Date, ByVal CodArt As String, ByVal DesArt As String, ByVal Qta As Double) As Boolean
        Try
            Dim TmpCS As String = String.Empty
            TmpCS = GetStringaConnessioneDBExt()
            If Not String.IsNullOrEmpty(TmpCS) Then SqlConn = New SqlConnection(GetStringaConnessioneDBExt)
            Dim StrSql As String = ""
            Dim NRecordInsert As Integer = 0
            Dim StrRfid As String = String.Empty
            Dim StrRfidHEX As String = String.Empty
            'StrSql = "INSERT INTO [Etichette] ([Codditt_Etichetta],[Tipork_Etichetta],[Anno_Etichetta],[Serie_Etichetta],[Numdoc_Etichetta],[Riga_Etichetta],[DataDoc_Etichetta],[Data_Etichetta],[CodArt_Etichetta],[DesArt_Etichetta],[Qta_Etichetta],[TagRfid_Etichetta],[TagRfidHEX_Etichetta]) "
            StrSql = "INSERT INTO [Etichette] ([Codditt_Etichetta],[Tipork_Etichetta],[Anno_Etichetta],[Serie_Etichetta],[Numdoc_Etichetta],[Riga_Etichetta],[DataDoc_Etichetta],[CodArt_Etichetta],[DesArt_Etichetta],[Qta_Etichetta],[TagRfid_Etichetta],[TagRfidHEX_Etichetta]) "
            StrSql += "VALUES "
            StrSql += "(" & CStrSQL(oApp.Ditta)
            StrSql += "," & CStrSQL(strTipoDoc)
            StrSql += "," & nAnnoDoc
            StrSql += "," & CStrSQL(strSerieDoc)
            StrSql += "," & lNumDoc
            StrSql += "," & nRigaDoc
            'StrSql += "," & CDataSQL(DataDoc)
            StrSql += ",'" & DataDoc.ToString("yyyy-dd-MM") & "'"
            'StrSql += "," & CDataSQL(Now.Date)
            StrSql += "," & CStrSQL(CodArt)
            StrSql += "," & CStrSQL(DesArt)
            StrSql += "," & CDblSQL(Qta)
            'Info da inserire nel campo "TagRfid_Etichetta"
            'EAN13 (13 caratteri) + Anno Doc (yy) (2 caratteri) + N. doc (00000) (5 caratteri) + Serie doc (1 carattere) + N° riga doc (3 caratteri)
            '-----------------------------------------------
            StrRfid = IIf(CodArt.Length > 13, CodArt.Substring(0, 13), CodArt).ToString
            StrRfid += DataDoc.ToString("yy")
            StrRfid += lNumDoc.ToString("00000")
            StrRfid += IIf(String.IsNullOrEmpty(strSerieDoc) Or String.IsNullOrWhiteSpace(strSerieDoc), "0", strSerieDoc).ToString
            StrRfid += nRigaDoc.ToString("000")
            '-----------------------------------------------
            StrSql += "," & CStrSQL(StrRfid)
            'Converto la mia stringa di 24 numeri in esadecimale per passarli al DB che li passerà all'RFID (08-06-21)
            StrRfidHEX = IntegerToHex(StrRfid)
            StrSql += "," & CStrSQL(StrRfidHEX)
            StrSql += ");"
            'oApp.MsgBoxInfo(StrSql)
            'InputBox("", "", StrSql)
            SqlComm = New SqlCommand(StrSql, SqlConn)
            SqlConn.Open()
            NRecordInsert = SqlComm.ExecuteNonQuery()
            SqlConn.Close()
            If NRecordInsert > 0 Then Return True
            Return False
        Catch ex As Exception
            Dim strErr As String = String.Empty
            GestError(ex, Me, "", oApp.InfoError, oApp.ErrorLogFile, True)
            Return False
        End Try
    End Function
    Public Function IntegerToHex(ByVal Numero As String) As String
        Try
            Dim esadecimale As String = String.Empty
            Dim number As String = String.Empty
            Dim resto As String = String.Empty
            While Numero.Length <> 0
                divisione(Numero, 16, resto, number)
                Select Case resto
                    Case "10"
                        esadecimale = "A" & esadecimale
                    Case "11"
                        esadecimale = "B" & esadecimale
                    Case "12"
                        esadecimale = "C" & esadecimale
                    Case "13"
                        esadecimale = "D" & esadecimale
                    Case "14"
                        esadecimale = "E" & esadecimale
                    Case "15"
                        esadecimale = "F" & esadecimale
                    Case Else
                        esadecimale = resto & esadecimale
                End Select
                Numero = number
                resto = String.Empty
                number = String.Empty
            End While
            Return esadecimale
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
            Return String.Empty
        End Try
    End Function
    Public Sub divisione(ByVal numeratore As String, denominatore As Integer, ByRef resto As String, ByRef quoziente As String)
        Try
            Dim numero As Integer
            While numeratore.Length <> 0
                Dim strTmp As String = CStr(IIf(String.IsNullOrEmpty(resto) OrElse resto = "0", "", resto)) & numeratore.Substring(0, 1)
                numero = CInt(strTmp)
                quoziente &= numero \ denominatore
                If quoziente = "0" Then quoziente = String.Empty
                resto = CType(numero Mod denominatore, String)
                numeratore = numeratore.Substring(1, numeratore.Length - 1)
            End While
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
        End Try
    End Sub
    Public Overridable Function CancellaEtichetteDBExt(ByVal strTipoDoc As String, ByVal nAnnoDoc As Integer, ByVal strSerieDoc As String, ByVal lNumDoc As Integer, ByVal nRigaDoc As Integer) As Boolean
        Try
            Dim TmpCS As String = String.Empty
            TmpCS = GetStringaConnessioneDBExt()
            If Not String.IsNullOrEmpty(TmpCS) Then SqlConn = New SqlConnection(GetStringaConnessioneDBExt)
            Dim StrSql As String = ""
            Dim NRecordDelete As Integer = 0
            StrSql = "DELETE FROM [Etichette] "
            StrSql += "WHERE "
            StrSql += "[Codditt_Etichetta] = " & CStrSQL(oApp.Ditta)
            StrSql += " And [Tipork_Etichetta] = " & CStrSQL(strTipoDoc)
            StrSql += " And [Anno_Etichetta] = " & nAnnoDoc
            StrSql += " And [Serie_Etichetta] = " & CStrSQL(strSerieDoc)
            StrSql += " And [Numdoc_Etichetta] = " & lNumDoc
            If nRigaDoc > 0 Then StrSql += " And [Riga_Etichetta] = " & nRigaDoc
            StrSql += ";"
            SqlComm = New SqlCommand(StrSql, SqlConn)
            SqlConn.Open()
            NRecordDelete = SqlComm.ExecuteNonQuery()
            SqlConn.Close()
            If NRecordDelete > 0 Then Return True
            Return False
        Catch ex As Exception
            Dim strErr As String = String.Empty
            GestError(ex, Me, "", oApp.InfoError, oApp.ErrorLogFile, True)
            Return False
        End Try
    End Function
    Public Overridable Function CancellaEtichettePrecedentiADataDBExt(ByVal DataRifDelete As Date, ByVal QualeDataRif As String, ByRef NRecordDelete As Integer) As Boolean
        Try
            'Legenda QualeDataRif: D=Documento ; S=Stampa
            Dim TmpCS As String = String.Empty
            TmpCS = GetStringaConnessioneDBExt()
            If Not String.IsNullOrEmpty(TmpCS) Then SqlConn = New SqlConnection(GetStringaConnessioneDBExt)
            Dim StrSql As String = ""
            NRecordDelete = 0
            StrSql = "DELETE FROM [Etichette] "
            StrSql += "WHERE "
            If QualeDataRif = "D" Then StrSql += "[DataDoc_Etichetta] < '" & DataRifDelete.ToString("yyyy-dd-MM") & "'"
            If QualeDataRif = "S" Then StrSql += "[Data_Etichetta] < '" & DataRifDelete.ToString("yyyy-dd-MM") & "'"
            'If QualeDataRif = "D" Then StrSql += "[DataDoc_Etichetta] < " & CDataSQL(DataRifDelete)
            'If QualeDataRif = "S" Then StrSql += "[Data_Etichetta] < " & CDataSQL(DataRifDelete)
            StrSql += ";"
            SqlComm = New SqlCommand(StrSql, SqlConn)
            SqlConn.Open()
            NRecordDelete = SqlComm.ExecuteNonQuery()
            SqlConn.Close()
            If NRecordDelete > 0 Then Return True
            Return False
        Catch ex As Exception
            Dim strErr As String = String.Empty
            GestError(ex, Me, "", oApp.InfoError, oApp.ErrorLogFile, True)
            Return False
        End Try
    End Function
    Public Overridable Function CancellaEtichetteAllDBExt(ByRef NRecordDelete As Integer) As Boolean
        Try
            Dim TmpCS As String = String.Empty
            TmpCS = GetStringaConnessioneDBExt()
            If Not String.IsNullOrEmpty(TmpCS) Then SqlConn = New SqlConnection(GetStringaConnessioneDBExt)
            Dim StrSql As String = ""
            NRecordDelete = 0
            StrSql = "DELETE FROM [Etichette];"
            SqlComm = New SqlCommand(StrSql, SqlConn)
            SqlConn.Open()
            NRecordDelete = SqlComm.ExecuteNonQuery()
            SqlConn.Close()
            If NRecordDelete > 0 Then Return True
            Return False
        Catch ex As Exception
            Dim strErr As String = String.Empty
            GestError(ex, Me, "", oApp.InfoError, oApp.ErrorLogFile, True)
            Return False
        End Try
    End Function
    Public Overridable Function SelezionaDaDocMagaRFID(ByVal strDitta As String, ByVal strTipodoc As String, ByVal lDaclie As Integer,
                                                     ByVal lAclie As Integer, ByVal dtDadata As Date, ByVal dtAdata As Date,
                                                     ByVal lDaanno As Integer, ByVal lAanno As Integer, ByVal strDaserie As String,
                                                     ByVal strAserie As String, ByVal lDanumero As Integer, ByVal lAnumero As Integer,
                                                     ByVal lDariga As Integer, ByVal lAriga As Integer, ByVal bDocMagazz As Boolean,
                                                     ByVal bColli As Boolean, ByRef dttOut As DataTable) As Boolean
        Dim strSQL As String = ""
        Dim strTableTesta As String
        Dim strTableCorpo As String
        Try
            If bDocMagazz = True Then
                'magazzino
                If strTipodoc = "W" Then
                    strTableTesta = "testprb"
                    strTableCorpo = "movprb"
                Else
                    strTableTesta = "testmag"
                    strTableCorpo = "movmag"
                End If

                'da documenti di magazzino
                strSQL = "Select mm_codart As Codart, " & NTSCStr(IIf(bColli, "mm_colli", "mm_quant")) & " As Quant, ar_codtagl As CodTagl, mm_tipork As Tipork, mm_anno As Anno, mm_serie As Serie, mm_numdoc As Numero, mm_riga As Riga, tm_datdoc As Dtdoc, tm_conto As Conto, ar_qtacon2 As Qtacon2, mm_ump As UmQuant, mm_unmis As UmColli, ar_misura1 As Misura1 " &
                  " FROM " & strTableTesta & " INNER JOIN (artico INNER JOIN " & strTableCorpo & " On (" & strTableCorpo & ".codditt = artico.codditt) And (artico.ar_codart = mm_codart)) On (" & strTableTesta & ".codditt = " & strTableCorpo & ".codditt) And (tm_tipork = mm_tipork) And (tm_serie = mm_serie) And (tm_anno = mm_anno) And (tm_numdoc = mm_numdoc)" &
                  " WHERE " & strTableCorpo & ".codditt = " & CStrSQL(strDitta) &
                  " And mm_tipork = " & CStrSQL(strTipodoc) &
                  " And tm_conto >= " & lDaclie &
                  " And tm_conto <= " & lAclie &
                  " And tm_datdoc >= " & CDataSQL(dtDadata) &
                  " And tm_datdoc <= " & CDataSQL(dtAdata) &
                  " And mm_anno >= " & lDaanno &
                  " And mm_anno <= " & lAanno &
                  " And mm_serie >= " & CStrSQL(strDaserie) &
                  " And mm_serie <= " & CStrSQL(strAserie) &
                  " And mm_numdoc >= " & lDanumero &
                  " And mm_numdoc <= " & lAnumero &
                  " And mm_riga >= " & lDariga &
                  " And mm_riga <= " & lAriga
            Else
                'da ordini
                strSQL = "Select mo_codart As Codart," &
                  " " & NTSCStr(IIf(bColli, "mo_colli", "mo_quant")) & " As Quant," &
                  " ar_codtagl As CodTagl, mo_tipork As Tipork, mo_anno As Anno," &
                  " mo_serie As Serie, mo_numord As Numero, mo_riga As Riga," &
                  " td_datord As Dtdoc, td_conto As Conto, ar_qtacon2 As Qtacon2, mo_ump As UmQuant, mo_unmis As UmColli, ar_misura1 As Misura1 " &
                  " FROM testord INNER JOIN (artico INNER JOIN movord  On (movord.codditt = artico.codditt) And (artico.ar_codart = movord.mo_codart)) On (testord.codditt = movord.codditt) And (testord.td_tipork = movord.mo_tipork) And (testord.td_serie = movord.mo_serie) And (testord.td_anno = movord.mo_anno) And (testord.td_numord = movord.mo_numord)" &
                  " WHERE movord.codditt = " & CStrSQL(strDitta) &
                  " And mo_tipork = " & CStrSQL(strTipodoc) &
                  " And td_conto >= " & lDaclie &
                  " And td_conto <= " & lAclie &
                  " And td_datord >= " & CDataSQL(dtDadata) &
                  " And td_datord <= " & CDataSQL(dtAdata) &
                  " And mo_anno >= " & lDaanno &
                  " And mo_anno <= " & lAanno &
                  " And mo_serie >= " & CStrSQL(strDaserie) &
                  " And mo_serie <= " & CStrSQL(strAserie) &
                  " And mo_numord >= " & lDanumero &
                  " And mo_numord <= " & lAnumero &
                  " And mo_riga >= " & lDariga &
                  " And mo_riga <= " & lAriga &
                  " And mo_hhisetrfid = 'S'"
            End If

            dttOut = OpenRecordset(strSQL, CLE__APP.DBTIPO.DBAZI)
            Return True
        Catch ex As Exception
            '--------------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '--------------------------------------------------------------
        End Try
    End Function
#End Region

#Region "Dati QR etichette"

    Public Overridable Function OttieniDatiPerQR(ByVal strDitta As String,
                                                  ByVal lInstId As Integer) As DataTable
        Try
            Dim strSQL As String =
                "SELECT tt_hhweek, tt_anno, tt_hhprogr," &
                " caf_codarfo AS xx_codarfo, tt_numdoc, tt_riga, tt_serie, tt_tipork," &
                " td_hhulrif AS xx_hhulrif, bc_code AS xx_code" &
                " FROM TTSTMATR" &
                " LEFT JOIN MOVORD ON TTSTMATR.codditt = MOVORD.codditt" &
                " AND tt_tipork = mo_tipork AND tt_anno = mo_anno" &
                " AND tt_serie = mo_serie AND tt_numdoc = mo_numord AND tt_riga = mo_riga" &
                " LEFT JOIN TESTORD ON TTSTMATR.codditt = TESTORD.codditt" &
                " AND tt_tipork = td_tipork AND tt_anno = td_anno" &
                " AND tt_serie = td_serie AND tt_numdoc = td_numord" &
                " LEFT JOIN CODARFO ON TTSTMATR.codditt = CODARFO.codditt" &
                " AND tt_codart = caf_codart AND tt_conto = caf_conto" &
                " LEFT JOIN BARCODE ON TTSTMATR.codditt = BARCODE.codditt" &
                " AND tt_codart = bc_codart" &
                " WHERE TTSTMATR.codditt = " & CStrSQL(strDitta) &
                " AND instid = " & lInstId.ToString()

            Return OpenRecordset(strSQL, CLE__APP.DBTIPO.DBAZI, "TTSTMATR").Tables("TTSTMATR")
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
            Return Nothing
        End Try
    End Function

#End Region

End Class
