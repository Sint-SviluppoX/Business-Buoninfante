Imports System.Data
Imports System.Data.Common
Imports NTSInformatica.CLN__STD

Imports System
Public Class CLHORGSOR
  Inherits CLDORGSOR

  Public Overridable Function GetDestinazioneImport(ByVal strDitta As String,
                                                    ByVal nConto As Integer,
                                                    ByVal strCodDestEsterno As String) As DataTable
    Try
      Dim strSQL As String = "SELECT dd_coddest, dd_hhGiornoConsegna FROM DESTDIV" &
                             " WHERE codditt = " & CStrSQL(strDitta) &
                             " AND dd_conto = " & nConto.ToString() &
                             " AND dd_hhCodDestExc = " & CStrSQL(strCodDestEsterno)
      Return OpenRecordset(strSQL, CLE__APP.DBTIPO.DBAZI)
    Catch ex As Exception
      CLN__STD.GestErr(ex, Me, "")
      Return Nothing
    End Try
  End Function

End Class
