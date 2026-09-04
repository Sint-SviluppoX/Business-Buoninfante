Imports System.Windows.Forms
Imports NTSInformatica.CLN__STD

Partial Public Class BOORGSOR_ProgressView

  Public Overloads Function Init(ByRef Menu As CLE__MENU, ByRef Param As CLE__CLDP,
                                 Optional ByVal Ditta As String = "",
                                 Optional ByRef SharedControls As CLE__EVNT = Nothing) As Boolean
    Try
      oMenu = Menu
      oApp = oMenu.App
      DittaCorrente = If(Ditta <> "", Ditta, oApp.Ditta)
      GctlTipoDoc = ""
      MinimumSize = Size
      Return True
    Catch ex As Exception
      CLN__STD.GestErr(ex, Me, "")
      Return False
    End Try
  End Function

  Public Sub Aggiorna(ByVal messaggio As String, ByVal percentuale As Integer)
    Try
      lbProgress.Text = messaggio
      pbProgress.EditValue = Math.Max(0, Math.Min(100, percentuale))
      lbProgress.Refresh()
      pbProgress.Refresh()
      Application.DoEvents()
    Catch ex As Exception
      CLN__STD.GestErr(ex, Me, "")
    End Try
  End Sub

End Class
