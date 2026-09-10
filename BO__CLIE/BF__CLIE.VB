Imports System.Data
Imports NTSInformatica.CLN__STD
Imports System.Globalization


Imports System
Public Class CLF__CLIE
  Inherits CLE__CLIE

  Public _oClhClie As CLH__CLIE
  Public Property oClhClie() As CLH__CLIE
    Get
      If _oClhClie Is Nothing Then _oClhClie = CType(ocldClie, CLH__CLIE)
      Return _oClhClie
    End Get
    Set(ByVal value As CLH__CLIE)
      _oClhClie = value
    End Set
  End Property

#Region "Controllo codice destinazione Excel"

  Public Overrides Function DesgTestPreSalva() As Boolean
    Try
      NormalizzaCampiDestinazione()

      If Not MyBase.DesgTestPreSalva() Then Return False

      Dim righe() As DataRow = dttDestdivEx.Select(
        Nothing, Nothing, DataViewRowState.Added Or DataViewRowState.ModifiedCurrent)

      For Each riga As DataRow In righe
        If riga.IsNull("dd_hhCodDestExc") Then Continue For

        Dim codiceEsterno As Integer = NTSCInt(riga!dd_hhCodDestExc)
        If codiceEsterno <= 0 Then Continue For

        Dim locale As DataRow = TrovaDuplicatoLocale(riga, codiceEsterno)
        If locale IsNot Nothing Then
          SegnalaDuplicato(codiceEsterno, NTSCStr(locale!codditt),
                           NTSCInt(locale!dd_conto), NTSCInt(locale!dd_coddest))
          Return False
        End If

        Dim trovato As DataTable = oClhClie.GetDestinazioneConCodiceEsterno(
          codiceEsterno, NTSCStr(riga!codditt), NTSCInt(riga!dd_conto), NTSCInt(riga!dd_coddest))
        If trovato IsNot Nothing AndAlso trovato.Rows.Count > 0 Then
          SegnalaDuplicato(codiceEsterno, NTSCStr(trovato.Rows(0)!codditt),
                           NTSCInt(trovato.Rows(0)!dd_conto), NTSCInt(trovato.Rows(0)!dd_coddest))
          Return False
        End If
      Next

      Return True
    Catch ex As Exception
      CLN__STD.GestErr(ex, Me, "")
      Return False
    End Try
  End Function

  Private Sub NormalizzaCampiDestinazione()
    Try
      Dim righe() As DataRow = dttDestdivEx.Select(
        Nothing, Nothing, DataViewRowState.Added Or DataViewRowState.ModifiedCurrent)

      For Each riga As DataRow In righe
        If Not riga.IsNull("dd_hhCodDestExc") AndAlso
           NTSCInt(riga!dd_hhCodDestExc) <= 0 Then
          riga!dd_hhCodDestExc = DBNull.Value
        End If

        If Not riga.IsNull("dd_hhGiornoConsegna") Then
          Dim giorno As String = NTSCStr(riga!dd_hhGiornoConsegna).Trim()
          If giorno = "" OrElse giorno = "0" Then
            riga!dd_hhGiornoConsegna = DBNull.Value
          End If
        End If
      Next
    Catch ex As Exception
      CLN__STD.GestErr(ex, Me, "")
      Throw
    End Try
  End Sub

  Private Function TrovaDuplicatoLocale(ByVal corrente As DataRow,
                                        ByVal codiceEsterno As Integer) As DataRow
    Try
      For Each riga As DataRow In dttDestdivEx.Rows
        If riga.RowState = DataRowState.Deleted OrElse Object.ReferenceEquals(riga, corrente) Then Continue For
        If Not riga.IsNull("dd_hhCodDestExc") AndAlso
           NTSCInt(riga!dd_hhCodDestExc) = codiceEsterno Then Return riga
      Next
      Return Nothing
    Catch ex As Exception
      CLN__STD.GestErr(ex, Me, "")
      Return Nothing
    End Try
  End Function

  Private Sub SegnalaDuplicato(ByVal codiceEsterno As Integer, ByVal ditta As String,
                               ByVal conto As Integer, ByVal codDest As Integer)
    Try
      ThrowRemoteEvent(New NTSEventArgs("", "Il codice destinazione Excel " &
        codiceEsterno.ToString() & " è già associato alla destinazione " &
        codDest.ToString() & " del conto " & conto.ToString() &
        " (ditta " & ditta.Trim() & ")." & vbCrLf &
        "Inserire un codice diverso."))
    Catch ex As Exception
      CLN__STD.GestErr(ex, Me, "")
    End Try
  End Sub

#End Region

End Class
