Imports System.Data
Imports NTSInformatica.CLN__STD

Public Class FRO__DESG
  Inherits FRM__DESG

#Region "Controlli importazione impegni"

  Private fm_hhImpExc As NTSGroupBox
  Private lb_hhCodDestExc As NTSLabel
  Private WithEvents ed_hhCodDestExc As NTSTextBoxNum
  Private lb_hhGiornoConsegna As NTSLabel
  Private WithEvents cb_hhGiornoConsegna As NTSComboBox

  Public Overrides Sub InitializeComponent()
    Try
      MyBase.InitializeComponent()

            fm_hhImpExc = New NTSGroupBox With {
        .Name = "fm_hhImpExc",
        .Text = "IMPORT EXCEL",
        .Left = 4,
        .Top = 516,
        .Width = 456,
        .Height = 120,
        .Tile = True
      }

            lb_hhCodDestExc = CreaEtichetta("lb_hhCodDestExc", "Codice destinazione excel", 4, 28, 132)
            ed_hhCodDestExc = New NTSTextBoxNum With {
        .Name = "ed_hhCodDestExc",
        .Left = 144,
        .Top = 28,
        .Width = 40
      }

            lb_hhGiornoConsegna = CreaEtichetta("lb_hhGiornoConsegna", "Giorno consegna", 4, 56, 88)
            cb_hhGiornoConsegna = New NTSComboBox With {
        .Name = "cb_hhGiornoConsegna",
        .Left = 100,
        .Top = 56,
        .Width = 100
      }

            fm_hhImpExc.Controls.AddRange(New Control() {
        lb_hhCodDestExc, ed_hhCodDestExc, lb_hhGiornoConsegna, cb_hhGiornoConsegna
      })

            'L'editor NTS registra il gruppo nella pagina "Altri dati".
            flAltriDati.Controls.Add(fm_hhImpExc)
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
    End Try
  End Sub

  Private Function CreaEtichetta(ByVal nome As String, ByVal testo As String,
                                 ByVal x As Integer, ByVal y As Integer,
                                 ByVal larghezza As Integer) As NTSLabel
    Try
      Return New NTSLabel With {
        .Name = nome,
        .Text = testo,
        .Left = x,
        .Top = y,
        .Width = larghezza,
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

            ed_hhCodDestExc.NTSSetParam(oMenu, "Codice destinazione Excel", "0", 9, 0, 999999999)
            cb_hhGiornoConsegna.NTSSetParam("Giorno di consegna")

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

            cb_hhGiornoConsegna.DataSource = giorni
            cb_hhGiornoConsegna.ValueMember = "codice"
            cb_hhGiornoConsegna.DisplayMember = "descrizione"
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
        End Try
    End Sub

    Public Overrides Sub Bindcontrols()
        Try
            MyBase.Bindcontrols()

            ed_hhCodDestExc.NTSDbField = "DESTDIV.dd_hhCodDestExc"
            cb_hhGiornoConsegna.NTSDbField = "DESTDIV.dd_hhGiornoConsegna"
            NTSFormAddDataBinding(dcDesg, fm_hhImpExc)
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
        End Try
    End Sub

    Protected Overrides Sub OnShown(ByVal e As EventArgs)
        Try
            MyBase.OnShown(e)

            flAltriDati.SuspendLayout()

            flAltriDati.Controls.SetChildIndex(
              fm_hhImpExc,
              flAltriDati.Controls.Count - 1
            )

            flAltriDati.ResumeLayout(True)

        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
        End Try
    End Sub

#End Region

End Class
