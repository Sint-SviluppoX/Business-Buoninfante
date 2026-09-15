Imports System.Data
Imports System.Data.Common
Imports System.Collections.Generic
Imports NTSInformatica.CLN__STD

Imports System
Public Class CLHORGSOR
  Inherits CLDORGSOR

  Public Overridable Function GetPostiLettoArticoli(ByVal strDitta As String,
                                                    ByVal codici As List(Of String)) As DataTable
    Try
      If codici Is Nothing OrElse codici.Count = 0 Then Return New DataTable()
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
