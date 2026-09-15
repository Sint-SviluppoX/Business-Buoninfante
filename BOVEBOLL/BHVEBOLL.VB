Imports System.Data
Imports System.Data.Common
Imports System.Collections.Generic
Imports NTSInformatica.CLN__STD

Imports System
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop.Excel
Imports System.Globalization

Public Class CLHVEBOLL
    Inherits CLDVEBOLL
    Public Overridable Function GetPostiLettoArticoli(ByVal strDitta As String,
                                                       ByVal codici As List(Of String)) As System.Data.DataTable
        Try
            If codici Is Nothing OrElse codici.Count = 0 Then Return New System.Data.DataTable()
            Dim valori As New List(Of String)()
            For Each codice As String In codici
                valori.Add(CStrSQL(codice))
            Next
            Dim strSQL As String = "SELECT ar_codart, ar_hhPostiLetto FROM ARTICO" &
                                   " WHERE codditt = " & CStrSQL(strDitta) &
                                   " AND ar_codart IN (" & String.Join(",", valori.ToArray()) & ")"
            Return OpenRecordset(strSQL, CLE__APP.DBTIPO.DBAZI)
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
            Return Nothing
        End Try
    End Function
    Dim SqlConn As SqlConnection = Nothing
    Dim SqlComm As SqlCommand = Nothing
    Dim SqlDA As SqlDataAdapter = Nothing
    Public Overridable Function GetStringaConnessioneDBExt() As String
        GetStringaConnessioneDBExt = String.Empty
        Dim strSQL As String = ""
        Dim dttOut As New Data.DataTable
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
    Public Function AggiornaCancellaRiferimentiDoc(ByVal strTipoDoc As String, ByVal nAnnoDoc As Integer, ByVal strSerieDoc As String, ByVal lNumDoc As Integer, ByVal nRigaDoc As Integer) As Boolean
        Try
            Dim StrSql As String = ""
            Dim NRecordDelete As Integer = 0
            StrSql = "UPDATE [ProduzioneBus] SET [DeleteDoc_ProduzioneBus] = 'S' "
            StrSql += "WHERE [DeleteDoc_ProduzioneBus] = 'N' AND "
            StrSql += "[Codditt_ProduzioneBus] = " & CStrSQL(oApp.Ditta)
            StrSql += " AND [Tipork_ProduzioneBus] = " & CStrSQL(strTipoDoc)
            StrSql += " AND [Anno_ProduzioneBus] = " & nAnnoDoc
            StrSql += " AND [Serie_ProduzioneBus] = " & CStrSQL(strSerieDoc)
            StrSql += " AND [Numdoc_ProduzioneBus] = " & lNumDoc
            If nRigaDoc > 0 Then StrSql += " AND [NumRiga_ProduzioneBus] = " & nRigaDoc
            StrSql += ";"
            SqlConn = New SqlConnection(GetStringaConnessioneDBExt)
            SqlComm = New SqlCommand(StrSql, SqlConn)
            SqlConn.Open()
            NRecordDelete = SqlComm.ExecuteNonQuery()
            SqlConn.Close()
            If NRecordDelete > 0 Then Return True
            Return False
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
            Return False
        End Try
    End Function
    Public Function CancellaRiferimentiDoc(ByVal strTipoDoc As String, ByVal nAnnoDoc As Integer, ByVal strSerieDoc As String, ByVal lNumDoc As Integer, ByVal nRigaDoc As Integer) As Boolean
        Try
            Dim StrSql As String = ""
            Dim NRecordDelete As Integer = 0
            StrSql = "DELETE FROM [ProduzioneBus] "
            StrSql += "WHERE "
            StrSql += "[Codditt_ProduzioneBus] = " & CStrSQL(oApp.Ditta)
            StrSql += ",[Tipork_ProduzioneBus] = " & CStrSQL(strTipoDoc)
            StrSql += ",[Anno_ProduzioneBus] = " & nAnnoDoc
            StrSql += ",[Serie_ProduzioneBus] = " & CStrSQL(strSerieDoc)
            StrSql += ",[Numdoc_ProduzioneBus] = " & lNumDoc
            If nRigaDoc > 0 Then StrSql += ",[NumRiga_ProduzioneBus] = " & nRigaDoc
            StrSql += ";"
            SqlConn = New SqlConnection(GetStringaConnessioneDBExt)
            SqlComm = New SqlCommand(StrSql, SqlConn)
            SqlConn.Open()
            NRecordDelete = SqlComm.ExecuteNonQuery()
            SqlConn.Close()
            If NRecordDelete > 0 Then Return True
            Return False
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
            Return False
        End Try
    End Function
    Public Function CheckRFidDocGenerate(ByVal strTipoDoc As String, ByVal nAnnoDoc As Integer, ByVal strSerieDoc As String, ByVal lNumDoc As Integer) As Boolean
        Try
            'Verifico se ci sono righe prodotte associate al documento
            Dim StrSql As String = ""
            Dim TmpDT As Data.DataTable
            StrSql = "SELECT [ID_ProduzioneBus] FROM [ProduzioneBus] "
            StrSql += "WHERE [Codditt_ProduzioneBus] = " & CStrSQL(oApp.Ditta)
            StrSql += " AND [Tipork_ProduzioneBus] = " & CStrSQL(strTipoDoc)
            StrSql += " AND [Anno_ProduzioneBus] = " & nAnnoDoc
            StrSql += " AND [Serie_ProduzioneBus] = " & CStrSQL(strSerieDoc)
            StrSql += " AND [Numdoc_ProduzioneBus] = " & lNumDoc
            StrSql += " AND [DeleteDoc_ProduzioneBus] = 'N'"
            CheckRFidDocGenerate = False
            SqlConn = New SqlConnection(GetStringaConnessioneDBExt)
            SqlComm = New SqlCommand(StrSql, SqlConn)
            SqlDA = New SqlDataAdapter(SqlComm)
            TmpDT = New Data.DataTable
            SqlDA.Fill(TmpDT)
            'Restituisco True solo se c'è almeno una riga
            If TmpDT.Rows.Count > 0 Then Return True
            Return False
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
            Return False
        End Try
    End Function


    Public dtnExcelDDT As Dictionary(Of String, String)
    Public ExcelDDT_SepDec As String = ""
    Public ExcelDDT_SepVal As String = ""
    Public ExcelDDT_SavePath As String = ""
    Public ExcelDDT_Extension As String = ""

    Public Overridable Function getDictionaryDDT() As Dictionary(Of String, String)
        Dim i As Integer = 0
        Try
            dtnExcelDDT = New Dictionary(Of String, String)
            i += 1
            dtnExcelDDT.Add("F" & i, "tm_riferim")
            i += 1
            dtnExcelDDT.Add("F" & i, "bc_code")
            i += 1
            dtnExcelDDT.Add("F" & i, "mm_descr")
            i += 1
            dtnExcelDDT.Add("F" & i, "caf_codarfo")
            i += 1
            dtnExcelDDT.Add("F" & i, "CONCAT(ar_descr, ' ', ar_desint) as xx_descr")
            i += 1
            dtnExcelDDT.Add("F" & i, "0 AS xx_raggruppamento")
            i += 1
            dtnExcelDDT.Add("F" & i, "0 AS xx_qtlineare")
            i += 1
            dtnExcelDDT.Add("F" & i, "0 AS mm_quant")

            Return dtnExcelDDT
        Catch ex As Exception
            '---------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '---------------------------------------------------------
        End Try
        Return Nothing
    End Function
    Public Overridable Function generaExcelDDT(ByVal strDitta As String, ByVal lNumdoc As Integer, ByVal strSerie As String, ByVal nAnno As Integer, ByVal strTipork As String) As Boolean
        Dim strSQL As String = String.Empty
        Dim dttTmp As Data.DataTable
        Try
            '*** Compongo la select ***
            strSQL = "SELECT "
            For Each strValue As String In dtnExcelDDT.Values
                strSQL &= $"{strValue},"
            Next
            strSQL = strSQL.TrimEnd(","c) & vbCrLf

            '*** Unisco le tabelle ***
            strSQL &= " FROM testmag " & vbCrLf
            strSQL &= " LEFT OUTER JOIN movmag ON movmag.codditt = testmag.codditt AND movmag.mm_numdoc = testmag.tm_numdoc AND movmag.mm_serie = testmag.tm_serie AND movmag.mm_anno = testmag.tm_anno AND movmag.mm_tipork = testmag.tm_tipork" & vbCrLf
            strSQL &= " LEFT OUTER JOIN testord ON movmag.codditt = testord.codditt AND movmag.mm_ornum = testord.td_numord AND movmag.mm_orserie = testord.td_serie AND movmag.mm_oranno = testord.td_anno AND movmag.mm_ortipo = testord.td_tipork " & vbCrLf
            strSQL &= " LEFT OUTER JOIN movord ON movmag.codditt = movord.codditt AND movmag.mm_orriga = movord.mo_riga AND movmag.mm_ornum = movord.mo_numord AND movmag.mm_orserie = movord.mo_serie AND movmag.mm_oranno = movord.mo_anno AND movmag.mm_ortipo = movord.mo_tipork " & vbCrLf
            strSQL &= " LEFT OUTER JOIN artico ON artico.codditt = movmag.codditt AND artico.ar_codart = movmag.mm_codart " & vbCrLf
            strSQL &= " LEFT OUTER JOIN anagra ON anagra.codditt = testmag.codditt AND anagra.an_conto = testmag.tm_conto " & vbCrLf
            strSQL &= " LEFT OUTER JOIN barcode ON barcode.codditt = movmag.codditt AND barcode.bc_codart = movmag.mm_codart " & vbCrLf
            strSQL &= " LEFT OUTER JOIN codarfo ON codarfo.codditt = artico.codditt AND codarfo.caf_codart = artico.ar_codart AND caf_conto = anagra.an_conto" & vbCrLf

            '*** Compongo la Where ***
            strSQL &= $" WHERE testmag.codditt = {CStrSQL(strDitta)}" & vbCrLf &
                $" AND testmag.tm_numdoc = {CStrSQL(lNumdoc)}" & vbCrLf &
                $" AND testmag.tm_serie = {CStrSQL(strSerie)}" & vbCrLf &
                $" AND testmag.tm_anno = {CStrSQL(nAnno)}" & vbCrLf &
                $" AND testmag.tm_tipork = {CStrSQL(strTipork)}"

            '*** Leggo i dati ***
            dttTmp = OpenRecordset(strSQL, CLE__APP.DBTIPO.DBAZI)

            '*** Esporta in Excel ***
            If dttTmp Is Nothing OrElse dttTmp.Rows.Count = 0 Then Return False

            If ExcelDDT_Extension.ToUpper = "CSV" Then
                ExportToCSV(dttTmp, $"DDT_{dttTmp.Rows(0)!tm_riferim}.csv")
            ElseIf ExcelDDT_Extension.ToUpper = "XLSX" Then
                ExportToExcel(dttTmp, $"DDT_{dttTmp.Rows(0)!tm_riferim}.xlsx")
            Else
                oApp.MsgBoxInfo($"Estensione '{ExcelDDT_Extension}' non supportata!")
                Return False
            End If

            Return True
        Catch ex As Exception
            '---------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '---------------------------------------------------------
        End Try
    End Function
    Public Overridable Sub ExportToExcel(ByVal dataTable As Data.DataTable, ByVal strNomeFile As String)
        Dim localPath As String
        Try
            Dim excelApp As New Application
            Dim workbook As Workbook = excelApp.Workbooks.Add()
            Dim worksheet As Worksheet = CType(workbook.Worksheets(1), Worksheet)

            ' Aggiungi i dati del DataTable (senza intestazioni delle colonne)
            For rowIndex As Integer = 0 To dataTable.Rows.Count - 1
                For colIndex As Integer = 0 To dataTable.Columns.Count - 1
                    Dim cellValue As String = dataTable.Rows(rowIndex)(colIndex).ToString()

                    ' Verifica se il valore è di tipo numerico o decimale
                    Dim decimalValue As Decimal
                    If Decimal.TryParse(cellValue, decimalValue) Then
                        ' Usa il separatore decimale personalizzato
                        cellValue = decimalValue.ToString(CultureInfo.InvariantCulture).Replace(".", ExcelDDT_SepDec)
                    End If

                    ' Scrivi il valore nella cella Excel
                    worksheet.Cells(rowIndex + 1, colIndex + 1) = cellValue
                Next
            Next

            ' Salva il file Excel
            If String.IsNullOrEmpty(ExcelDDT_SavePath) Then
                Dim desktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                localPath = System.IO.Path.Combine(desktopPath, strNomeFile)
            Else
                localPath = System.IO.Path.Combine(ExcelDDT_SavePath, strNomeFile)
            End If

            ' Salva il file Excel
            workbook.SaveAs(localPath)
            workbook.Close()
            excelApp.Quit()

            ' Rilascia gli oggetti COM
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet)
            worksheet = Nothing
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            workbook = Nothing
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp)
            excelApp = Nothing

        Catch ex As Exception
            ' Gestisci l'errore
            CLN__STD.GestErr(ex, Me, "")
        Finally
            GC.Collect()
        End Try
    End Sub
    Public Overridable Sub ExportToCSV(ByVal dataTable As Data.DataTable, ByVal strNomeFile As String)
        Dim localPath As String
        Try
            ' Definisci il percorso di salvataggio
            If String.IsNullOrEmpty(ExcelDDT_SavePath) Then
                Dim desktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                localPath = System.IO.Path.Combine(desktopPath, strNomeFile)
            Else
                localPath = System.IO.Path.Combine(ExcelDDT_SavePath, strNomeFile)
            End If

            ' Usa StreamWriter per scrivere il file CSV
            Using writer As New System.IO.StreamWriter(localPath, False)

                ' Aggiungi i dati del DataTable
                For rowIndex As Integer = 0 To dataTable.Rows.Count - 1
                    For colIndex As Integer = 0 To dataTable.Columns.Count - 1
                        Dim cellValue As String = dataTable.Rows(rowIndex)(colIndex).ToString()

                        ' Verifica se il valore è di tipo numerico o decimale
                        Dim decimalValue As Decimal
                        If Decimal.TryParse(cellValue, decimalValue) Then
                            ' Sostituisci il separatore decimale predefinito con quello personalizzato
                            cellValue = decimalValue.ToString(CultureInfo.InvariantCulture).Replace(".", ExcelDDT_SepDec)
                        End If

                        ' Controlla se il valore della cella contiene caratteri problematici
                        If cellValue.Contains(";") OrElse cellValue.Contains("""") OrElse cellValue.Contains(vbCr) OrElse cellValue.Contains(vbLf) Then
                            ' Se contiene caratteri speciali, aggiungi le virgolette e raddoppia eventuali virgolette interne
                            cellValue = Chr(34) & cellValue.Replace("""", """""") & Chr(34)
                        End If

                        writer.Write(cellValue)

                        If colIndex < dataTable.Columns.Count - 1 Then
                            writer.Write(ExcelDDT_SepVal)
                        End If
                    Next
                    writer.WriteLine()
                Next
            End Using

        Catch ex As Exception
            ' Gestisci l'errore
            CLN__STD.GestErr(ex, Me, "")
        Finally
            GC.Collect()
        End Try
    End Sub

End Class
