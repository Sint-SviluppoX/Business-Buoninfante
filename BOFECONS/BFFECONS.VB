Imports System.Data
Imports NTSInformatica.CLN__STD
Imports System.Globalization


Imports System
Public Class CLFFECONS
    Inherits CLEFECONS

    Public _oClhCons As CLHFECONS
    Public Property oClhCons() As CLHFECONS
        Get
            If _oClhCons Is Nothing Then _oClhCons = CType(oCldCons, CLHFECONS)
            Return _oClhCons
        End Get
        Set(ByVal value As CLHFECONS)
            _oClhCons = value
        End Set
    End Property

    Public Overrides Function ImportFattura_DeterminaConto(ByRef dttTmBase As DataTable, bDocRicevuto As Boolean, ByRef lFtContoCF As Integer, ByRef strTipoContoNew As String, dtrDatiGenDoc As DataRow, strFile As String) As Boolean
        Try
            '*** Imposto il codice fiscale per gli stati esteri ***
            If dttTmBase IsNot Nothing AndAlso dttTmBase.Rows.Count > 0 Then
                For Each dtrFetestmag As DataRow In dttTmBase.Rows
                    If String.IsNullOrEmpty(NTSCStr(dtrFetestmag!tm_cocodfis)) AndAlso NTSCStr(dtrFetestmag!tm_coidpaese) <> "IT" Then
                        dtrFetestmag!tm_coidcodice = (NTSCStr(dtrFetestmag!tm_codenom) & NTSCStr(dtrFetestmag!tm_conome) & NTSCStr(dtrFetestmag!tm_cocognome)).Replace(" ", "").ToUpper.PadRight(14, "0"c).Substring(0, 14)
                    End If
                Next
            End If

            Return MyBase.ImportFattura_DeterminaConto(dttTmBase, bDocRicevuto, lFtContoCF, strTipoContoNew, dtrDatiGenDoc, strFile)

        Catch ex As Exception
            '---------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '---------------------------------------------------------
        End Try
    End Function
    Public Overrides Function CreaCliFor(bDocRicevuto As Boolean, dttFetestmag As DataTable, ByRef lConto As Integer, strTipoConto As String) As Boolean
        Try
            '*** Imposto il codice fiscale per gli stati esteri ***
            If dttFetestmag IsNot Nothing AndAlso dttFetestmag.Rows.Count > 0 Then
                For Each dtrFetestmag As DataRow In dttFetestmag.Rows
                    If String.IsNullOrEmpty(NTSCStr(dtrFetestmag!tm_cocodfis)) AndAlso NTSCStr(dtrFetestmag!tm_coidpaese) <> "IT" Then
                        dtrFetestmag!tm_coidcodice = (NTSCStr(dtrFetestmag!tm_codenom) & NTSCStr(dtrFetestmag!tm_conome) & NTSCStr(dtrFetestmag!tm_cocognome)).Replace(" ", "").ToUpper.PadRight(14, "0"c).Substring(0, 14)
                    End If
                Next
            End If

            '*** Eseguo lo standard ***
            Return MyBase.CreaCliFor(bDocRicevuto, dttFetestmag, lConto, strTipoConto)
        Catch ex As Exception
            '---------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '---------------------------------------------------------
        End Try
    End Function


End Class