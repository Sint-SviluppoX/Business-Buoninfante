Imports System.Data
Imports NTSInformatica.CLN__STD
Imports System.Globalization
Imports System
Public Class CLFORGSOR
  Inherits CLEORGSOR
  Public _oClhGsor As CLHORGSOR
  Public Property oClhGsor() As CLHORGSOR
    Get
      If _oClhGsor Is Nothing Then _oClhGsor = CType(ocldGsor, CLHORGSOR)
      Return _oClhGsor
    End Get
    Set(ByVal value As CLHORGSOR)
      _oClhGsor = value
    End Set
  End Property
  Public Overridable Function GetDestinazioneImport(ByVal strCodDestEsterno As String,
                                                    ByVal nConto As Integer) As DataTable
    Try
      Return oClhGsor.GetDestinazioneImport(strDittaCorrente, nConto, strCodDestEsterno)
    Catch ex As Exception
      CLN__STD.GestErr(ex, Me, "")
      Return Nothing
    End Try
  End Function

    Public Overrides Function AfterColUpdate_CORPO_ec_quant(sender As Object, e As DataColumnChangeEventArgs) As Boolean
        Try
            Dim Tmp As Boolean = MyBase.AfterColUpdate_CORPO_ec_quant(sender, e)
            'Calcolo il volume (Pers. Alfy 24-03-21)
            'If Not IsNothing(dttArti) Then e.Row!ec_hhvol = NTSCDec(e.ProposedValue) * NTSCDec(dttArti.Rows(0)!ar_volume)
            'Calcolo peso netto e lordo (Pers. Alfy 15-03-22)
            If dttArti IsNot Nothing AndAlso dttArti.Rows.Count > 0 AndAlso
         dttArti.Columns.Contains("ar_volume") AndAlso
         dttArti.Columns.Contains("ar_pesonet") AndAlso
         dttArti.Columns.Contains("ar_pesolor") Then
                'Calcolo il volume (Pers. Alfy 24-03-21)
                e.Row!ec_hhvol = NTSCDec(e.ProposedValue) * NTSCDec(dttArti.Rows(0)!ar_volume)
                'Calcolo peso netto e lordo (Pers. Alfy 15-03-22)
                e.Row!ec_hhpn = NTSCDec(e.ProposedValue) * NTSCDec(dttArti.Rows(0)!ar_pesonet)
                e.Row!ec_hhpl = NTSCDec(e.ProposedValue) * NTSCDec(dttArti.Rows(0)!ar_pesolor)
            End If
            Return Tmp
        Catch ex As Exception
            '--------------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '--------------------------------------------------------------
        End Try

    End Function
    Public Overridable Function CalcolaTotaleVolume() As Decimal
        Try
            Dim TmpTotaleVolume As Decimal = 0
            If Not IsNothing(dsShared) AndAlso dsShared.Tables.Contains("CORPO") Then
                TmpTotaleVolume = NTSCDec(dsShared.Tables("CORPO").Compute("SUM(ec_hhvol)", ""))
            End If
            Return ArrDbl(TmpTotaleVolume, oApp.NDecQta)
        Catch ex As Exception
            '--------------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '--------------------------------------------------------------	
        End Try
    End Function
    Public Overridable Function CalcolaTotalePesoNetto() As Decimal
        Try
            Dim TmpTotalePesoNetto As Decimal = 0
            If Not IsNothing(dsShared) AndAlso dsShared.Tables.Contains("CORPO") Then
                TmpTotalePesoNetto = NTSCDec(dsShared.Tables("CORPO").Compute("SUM(ec_hhpn)", ""))
            End If
            Return ArrDbl(TmpTotalePesoNetto, oApp.NDecQta)
        Catch ex As Exception
            '--------------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '--------------------------------------------------------------	
        End Try
    End Function
    Public Overridable Function CalcolaTotalePesoLordo() As Decimal
        Try
            Dim TmpTotalePesoLordo As Decimal = 0
            If Not IsNothing(dsShared) AndAlso dsShared.Tables.Contains("CORPO") Then
                TmpTotalePesoLordo = NTSCDec(dsShared.Tables("CORPO").Compute("SUM(ec_hhpl)", ""))
            End If
            Return ArrDbl(TmpTotalePesoLordo, oApp.NDecQta)
        Catch ex As Exception
            '--------------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '--------------------------------------------------------------	
        End Try
    End Function
    Public Overrides Function CalcolaTotali() As Boolean
        Try
            If dttET Is Nothing Then Return True
            If dttET.Rows.Count = 0 Then Return True
            Dim TmpTotVol As Decimal = 0
            Dim TmpTotPesoNetto As Decimal = 0
            Dim TmpTotPesoLordo As Decimal = 0
            TmpTotVol = CalcolaTotaleVolume()
            TmpTotPesoNetto = CalcolaTotalePesoNetto()
            TmpTotPesoLordo = CalcolaTotalePesoLordo()
            If TmpTotVol <> NTSCDec(dttET.Rows(0)!et_hhtotvol) Then dttET.Rows(0)!et_hhtotvol = CalcolaTotaleVolume()
            If TmpTotPesoNetto <> NTSCDec(dttET.Rows(0)!et_hhtotpn) Then dttET.Rows(0)!et_hhtotpn = CalcolaTotalePesoNetto()
            If TmpTotPesoLordo <> NTSCDec(dttET.Rows(0)!et_hhtotpl) Then dttET.Rows(0)!et_hhtotpl = CalcolaTotalePesoLordo()
            Return MyBase.CalcolaTotali()
        Catch ex As Exception
            '--------------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '--------------------------------------------------------------	
        End Try
    End Function
    Public Overrides Function ApriOrdine(strDitta As String, bNew As Boolean, strTipoDoc As String, nAnno As Integer, strSerie As String, lNumdoc As Integer, ByRef ds As DataSet) As Boolean
        Try
            Dim TmpRis As Boolean = False
            TmpRis = MyBase.ApriOrdine(strDitta, bNew, strTipoDoc, nAnno, strSerie, lNumdoc, ds)
            If TmpRis = True AndAlso ds.Tables.Contains("CORPO") AndAlso ds.Tables("CORPO").Columns.Contains("ec_hhisetrfid") Then
                'Se Ordine di Produzione metto a "S" il valore di default del campo pers. che indica se � una riga che dovr� andare nelle etichette rfid
                If strTipoDoc = "H" Then
                    ds.Tables("CORPO").Columns("ec_hhisetrfid").DefaultValue = "S"
                Else
                    ds.Tables("CORPO").Columns("ec_hhisetrfid").DefaultValue = "N"
                End If
            End If
            Return TmpRis
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
        End Try
    End Function

    'Public Overrides Sub SetDefaultValueMovord(ByRef ds As DataSet, bMovordImp As Boolean)
    '  Try
    '    MyBase.SetDefaultValueMovord(ds, bMovordImp)
    '    If ds.Tables.Contains("MOVORD") AndAlso ds.Tables("MOVORD").Columns.Contains("mo_hhisetrfid") Then
    '      ds.Tables("MOVORD").Columns("mo_hhisetrfid").DefaultValue
    '    End If
    '  Catch ex As Exception
    '    CLN__STD.GestErr(ex, Me, "")
    '  End Try
    'End Sub
End Class
