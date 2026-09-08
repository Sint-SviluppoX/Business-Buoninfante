Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.IO
Imports System.Linq
Imports System.Text
Imports NTSInformatica
Imports NTSInformatica.CLN__STD

Public Class CLD__HHID
    Inherits CLD__BASE

#Region "Variabili"
    Public FieldsList As Dictionary(Of String, String)
#End Region

#Region "Init"
    'Info Dictionary: Rappresenta una raccolta di chiavi e valori.
    'Dictionary<TKey,TValue>
    'TKey - (Tipo di chiavi nel dizionario)
    'TValue - (Tipo di valori nel dizionario)
    Public Overridable Function initFieldList(ByVal TipoFile As String) As Boolean
        'Legenda:
        'TipoFile : I (Impegni) - A (Articoli)

        'Variabile per identificare, con il dictionary, la posizione della colonna nel file 
        Dim iCol As Integer = 0
        Try
            '*** Inizializzo il dictionary ***
            FieldsList = New Dictionary(Of String, String)

            '*** Assegno le chiavi in base al tipo di tabella da leggere ***
            Select Case TipoFile
                Case "Amazon"
                    iCol += 1
                    FieldsList.Add("order-id".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("order-item-id".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("purchase-date".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("payments-date".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("buyer-email".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("buyer-name".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("buyer-phone-number".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("sku".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("number-of-items".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("product-name".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("quantity-purchased".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("currency".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("item-price".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("item-tax".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("shipping-price".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("shipping-tax".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("ship-service-level".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("recipient-name".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("ship-address-1".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("ship-address-2".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("ship-address-3".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("ship-city".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("ship-state".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("ship-postal-code".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("ship-country".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("ship-phone-number".ToUpper(), "F" & iCol)

                    iCol += 1
                    FieldsList.Add("item-promotion-discount".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("item-promotion-id".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("ship-promotion-discount".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("ship-promotion-id".ToUpper(), "F" & iCol)

                    iCol += 1
                    FieldsList.Add("delivery-start-date".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("delivery-end-date".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("delivery-time-zone".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("delivery-Instructions".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("earliest-ship-date".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("latest-ship-date".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("earliest-delivery-date".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("latest-delivery-date".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("is-business-order".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("purchase-order-number".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("price-designation".ToUpper(), "F" & iCol)

                    iCol += 1
                    FieldsList.Add("is-prime".ToUpper(), "F" & iCol)

                    iCol += 1
                    FieldsList.Add("buyer-company-name".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("is-iba".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("is-buyer-requested-cancellation".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("buyer-requested-cancel-reason".ToUpper(), "F" & iCol)
                Case "Xbio", "SleepH"
                    iCol += 1
                    FieldsList.Add("Total-Shipping-Tax-included".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("Total-Shipping-Tax-excluded".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("InvoiceTaxType-TaxID-TaxName-TaxAmount".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("Product-Unit-Price-Tax-Included".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("Product-Unit-Price-Tax-Excluded".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("Order-Reference".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("Order-Creation-Date".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("Customer-Email".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("Customer-Group-Names".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("Product-Name-With-Combination".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("Product-Quantity".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("Riferimento-del-prodotto".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("Combination-EAN13".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("Invoice-Address-Firstname".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("Invoice-Address-Lastname".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("Invoice-Address-VAT-Number".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("Invoice-Address-DNI".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("Invoice-Address1".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("Invoice-Address2".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("Invoice-Address-Postcode".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("Invoice-Address-City".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("Invoice-Address-State-Name".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("Invoice-Address-Country-Name".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("Invoice-Address-Phone".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("Delivery-Address-Firstname".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("Delivery-Address-Lastname".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("Delivery-Address-VAT-Number".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("Delivery-Address-DNI".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("Delivery-Address1".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("Delivery-Address2".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("Delivery-Address-Postcode".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("Delivery-Address-City".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("Delivery-Address-State-Name".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("Delivery-Address-Country-Name".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("Delivery-Address-Phone".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("Carrier-Name".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("Payment-Total-Discounts-Tax-excluded".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("Total-Discounts-Tax-included".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("Total-Discounts-Tax-excluded".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("Total-Paid-Tax-excluded".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("Total-Paid-Tax-included".ToUpper(), "F" & iCol)
                    iCol += 1
                    FieldsList.Add("Total-Paid".ToUpper(), "F" & iCol)
            End Select

            '***  datatable a Business datatable ***
            'Metto tutte le KEY del dictonary in una lista di stringhe e le ciclo per avere duplicare gli item al contratio (Valore=Key - Key=Valore)
            Dim KeysList As List(Of String) = FieldsList.Keys.ToList()
            For Each key As String In KeysList
                FieldsList.Add(FieldsList.Item(key), key)
            Next
        Catch ex As Exception
            '---------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '---------------------------------------------------------
        End Try
    End Function
#End Region

#Region "Funzioni Generiche"
    Public Overridable Function OpenConn(ByVal strPath As String) As OleDb.OleDbConnection
        Dim connectionString As String
        Dim dbConn As OleDb.OleDbConnection
        Try
            '*** Apertura della connessione ***
            connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strPath & ";Extended Properties='EXCEL 8.0;HDR=No;IMEX=1;'"
                    dbConn = New OleDb.OleDbConnection(connectionString)

            Return dbConn
        Catch ex As Exception
            '---------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            Return Nothing
            '---------------------------------------------------------
        End Try
    End Function
    Public Overridable Function LetturaExcel(ByVal strPath As String, ByVal tableName As String, ByVal strTipoFile As String, ByVal Optional strSelect As String = "*", ByVal Optional strWhere As String = "", ByVal Optional strGroup As String = "", ByVal Optional strOrder As String = "", ByVal Optional strNomeFoglioExcel As String = "", ByRef Optional dsOut As DataSet = Nothing) As Boolean
        Dim dbConn As OleDb.OleDbConnection
        Dim dbCmd As OleDb.OleDbDataAdapter

        Dim StrSQL As String = String.Empty
        Dim dsTmp As New DataSet

        Try
            '*** Apro la connessione ***
            dbConn = OpenConn(strPath)

            '*** Creazione della query ***
            StrSQL = "SELECT " & NTSCStr(IIf(String.IsNullOrEmpty(strSelect) OrElse String.IsNullOrEmpty(strSelect.Trim), "*", strSelect))

            '*** Specifico il foglio ***"
            StrSQL &= " FROM [" & NTSCStr(IIf(String.IsNullOrEmpty(strNomeFoglioExcel.Trim), "Foglio1", strNomeFoglioExcel)) & "$] "

            '**** Escludo tutte le righe che hanno la prima colonna vuota ***
            StrSQL += " WHERE F1 <>'' "

            '*** Raggruppo se necessario ***
            If Not String.IsNullOrEmpty(strGroup) AndAlso Not String.IsNullOrEmpty(strGroup.Trim) Then StrSQL &= " GROUP BY " & strGroup

            '*** Ordino se necessario ***
            If Not String.IsNullOrEmpty(strOrder) AndAlso Not String.IsNullOrEmpty(strOrder.Trim) Then StrSQL &= " ORDER BY " & strOrder

            '*** Esecuzione della query ***
            dbCmd = New OleDb.OleDbDataAdapter(StrSQL, dbConn)
            dbCmd.TableMappings.Add("Table", tableName)

            '*** Riempimento del dataset ***
            dbCmd.Fill(dsTmp)

            '*** Cancello il primo rigo con i nomi colonna ***
            Dim nRow As Integer = 0
            If strTipoFile = "Amazon" Then
                If tableName = "Testa" Or tableName = "Corpo" Then nRow = dsTmp.Tables(tableName).Rows.IndexOf(dsTmp.Tables(tableName).Select("F1 LIKE " & CStrSQL("%order-id%"))(0))
            ElseIf strTipoFile = "Xbio" Or strTipoFile = "SleepH" Then
                If tableName = "Testa" Or tableName = "Corpo" Then nRow = dsTmp.Tables(tableName).Rows.IndexOf(dsTmp.Tables(tableName).Select("F6 LIKE " & CStrSQL("%Order Reference%"))(0))
            Else
                Return False
            End If

            dsTmp.Tables(tableName).Rows(nRow).Delete()
            dsTmp.Tables(tableName).AcceptChanges()

            '*** Rinomino le colonne ***
            rinominaColonneDataTable(strTipoFile, dsTmp.Tables(tableName))

            '*** Controllo se il dataset può ricevere la tabella ***
            If dsOut Is Nothing Then dsOut = New DataSet
            If dsOut.Tables.Contains(tableName) Then dsOut.Tables.Remove(tableName)

            '*** Restituisco la tabella risultante ***
            dsOut.Tables.Add(dsTmp.Tables(tableName).Copy)

            '*** Chiusura della connessione ***
            dbConn.Close()
            Return True
        Catch ex As Exception
            '---------------------------------------------------------
            If ex.Message.Contains("Nessun valore specificato per alcuni parametri necessari.") Or ex.Message.Contains("Indice oltre i limiti della matrice") Then
                oApp.MsgBoxExclamation("File non conforme al tracciato dei dati da importare")
            Else
                CLN__STD.GestErr(ex, Me, "")
            End If
            Return False
            '---------------------------------------------------------
        End Try
    End Function
    Public Overridable Function rinominaColonneDataTable(ByVal TipoFile As String, ByRef dttOut As DataTable) As Boolean
        Try
            If FieldsList Is Nothing Then initFieldList(TipoFile)

            For Each dtc As DataColumn In dttOut.Columns
                dtc.ColumnName = FieldsList.Item(dtc.ColumnName)
            Next
            dttOut.AcceptChanges()

            Return True
        Catch ex As Exception
            '---------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            Return False
            '---------------------------------------------------------
        End Try
    End Function
    Public Overridable Function GetListinoType(ByVal strDitta As String, ByVal nListino As Integer) As Boolean
        Dim dttTmp As New DataTable
        Try
            If Not ValCodiceDb(NTSCStr(nListino), strDitta, "TABLIST", "N", "", dttTmp) Then Return False
            If dttTmp.Rows.Count = 0 Then Return False

            Return NTSCStr(dttTmp.Rows(0)!tb_ivato) = "S"
        Catch ex As Exception
            '---------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '---------------------------------------------------------
        End Try
    End Function
#End Region

#Region "utility"
    '*** Gestione log ***
    Public Overridable Sub ScriviLog(ByVal strText As String)
        Dim strFolderPath As String
        Dim strFileName As String
        Try
            '*** Inizializzo il percorso e il nome del file ***
            strFolderPath = oApp.AscDir & "\BS__HHID\"
            strFileName = "Log_" & Now.Date.ToString("ddMMyyyy") & ".txt"

            '*** Controllo se è già stata creata la cartella ***
            If Not IO.Directory.Exists(strFolderPath) Then
                Directory.CreateDirectory(strFolderPath)
            End If

            '*** Controllo se è già stato creato il file di log ***
            If Not IO.File.Exists(strFolderPath & strFileName) Then
                '*** Creo il file di log e lo chiudo ***
                Dim fs As IO.FileStream = File.Create(strFolderPath & strFileName)
                fs.Close()
            End If

            '*** Apro il file di log creato ***
            Dim LogFile As IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(strFolderPath & strFileName, True)

            '*** Aggiungo il log corretne alla lista ***
            LogFile.WriteLine("[" & Now.ToString & "] " & strText)

            '*** Chiudo il file di log ***
            LogFile.Close()
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
        End Try
    End Sub
#End Region
End Class
