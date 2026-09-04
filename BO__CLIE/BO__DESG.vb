Imports System.Data
Imports NTSInformatica.CLN__STD

Public Class FRO__DESG
  Inherits FRM__DESG

#Region "Controlli importazione impegni"

  Private fmhhImportazioneImpegni As NTSGroupBox
  Private lbhhCodDestExc As NTSLabel
  Private WithEvents edhhCodDestExc As NTSTextBoxNum
  Private lbhhGiornoConsegna As NTSLabel
  Private WithEvents cbhhGiornoConsegna As NTSComboBox

  Public Overrides Sub InitializeComponent()
    Try
      MyBase.InitializeComponent()

      fmhhImportazioneImpegni = New NTSGroupBox With {
        .Name = "fmhhImportazioneImpegni",
        .Text = "IMPORTAZIONE IMPEGNI",
        .Left = 4,
        .Top = 136,
        .Width = 448,
        .Height = 92,
        .Tile = False
      }

      lbhhCodDestExc = CreaEtichetta("lbhhCodDestExc", "Codice destinazione Excel", 4, 28)
      edhhCodDestExc = New NTSTextBoxNum With {
        .Name = "edhhCodDestExc",
        .Left = 172,
        .Top = 28,
        .Width = 272
      }

      lbhhGiornoConsegna = CreaEtichetta("lbhhGiornoConsegna", "Giorno di consegna", 4, 52)
      cbhhGiornoConsegna = New NTSComboBox With {
        .Name = "cbhhGiornoConsegna",
        .Left = 172,
        .Top = 52,
        .Width = 272
      }

      fmhhImportazioneImpegni.Controls.AddRange(New Control() {
        lbhhCodDestExc, edhhCodDestExc, lbhhGiornoConsegna, cbhhGiornoConsegna
      })

            'DATI PRINCIPALI 2 appartiene alla prima NTSTabPage di FRM__DESG.
            flPrincipale.Controls.Add(fmhhImportazioneImpegni)
            fmhhImportazioneImpegni.BringToFront()
    Catch ex As Exception
      CLN__STD.GestErr(ex, Me, "")
    End Try
  End Sub

  Private Function CreaEtichetta(ByVal nome As String, ByVal testo As String,
                                 ByVal x As Integer, ByVal y As Integer) As NTSLabel
    Try
      Return New NTSLabel With {
        .Name = nome,
        .Text = testo,
        .Left = x,
        .Top = y,
        .Width = 164,
        .Height = 20,
        .NTSBordeStyle = NTSLabel.NTSBorderStyle.FieldCaption,
        .UseMnemonic = False
      }
    Catch ex As Exception
      CLN__STD.GestErr(ex, Me, "")
      Return New NTSLabel()
    End Try
  End Function

  Public Overrides Sub InitControls()
    Try
      MyBase.InitControls()

      edhhCodDestExc.NTSSetParam(oMenu, "Codice destinazione Excel", "0", 9, 0, 999999999)
      cbhhGiornoConsegna.NTSSetParam("Giorno di consegna")

      Dim giorni As New DataTable
      giorni.Columns.Add("codice", GetType(String))
      giorni.Columns.Add("descrizione", GetType(String))
      giorni.Rows.Add("", "(Nessuno)")
      giorni.Rows.Add("Lun", "Lunedì")
      giorni.Rows.Add("Mar", "Martedì")
      giorni.Rows.Add("Mer", "Mercoledì")
      giorni.Rows.Add("Gio", "Giovedì")
      giorni.Rows.Add("Ven", "Venerdì")
      giorni.Rows.Add("Sab", "Sabato")
      giorni.Rows.Add("Dom", "Domenica")

      cbhhGiornoConsegna.DataSource = giorni
      cbhhGiornoConsegna.ValueMember = "codice"
      cbhhGiornoConsegna.DisplayMember = "descrizione"
    Catch ex As Exception
      CLN__STD.GestErr(ex, Me, "")
    End Try
  End Sub

  Public Overrides Sub Bindcontrols()
    Try
      MyBase.Bindcontrols()

      edhhCodDestExc.NTSDbField = "DESTDIV.dd_hhCodDestExc"
      cbhhGiornoConsegna.NTSDbField = "DESTDIV.dd_hhGiornoConsegna"
      NTSFormAddDataBinding(dcDesg, fmhhImportazioneImpegni)
    Catch ex As Exception
      CLN__STD.GestErr(ex, Me, "")
    End Try
  End Sub

#End Region

End Class
