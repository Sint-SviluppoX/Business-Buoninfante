Imports System.Data
Imports System.Data.Common
Imports NTSInformatica.CLN__STD

Imports System
Public Class CLH__CLIE
  Inherits CLD__CLIE

  Public Overridable Function GetDestinazioneConCodiceEsterno(ByVal nCodiceEsterno As Integer,
                                                               ByVal strDitta As String,
                                                               ByVal nConto As Integer,
                                                               ByVal nCodDest As Integer) As DataTable
    Try
      Dim strSQL As String = "SELECT TOP 1 codditt, dd_conto, dd_coddest FROM DESTDIV" &
                             " WHERE dd_hhCodDestExc = " & nCodiceEsterno.ToString() &
                             " AND NOT (codditt = " & CStrSQL(strDitta) &
                             " AND dd_conto = " & nConto.ToString() &
                             " AND dd_coddest = " & nCodDest.ToString() & ")" &
                             " ORDER BY codditt, dd_conto, dd_coddest"
      Return OpenRecordset(strSQL, CLE__APP.DBTIPO.DBAZI)
    Catch ex As Exception
      CLN__STD.GestErr(ex, Me, "")
      Return Nothing
    End Try
  End Function

End Class
