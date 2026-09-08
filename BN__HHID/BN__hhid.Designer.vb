Partial Public Class FRM__HHID
  Inherits FRM__CHIL

  <System.Diagnostics.DebuggerNonUserCode()>
  Public Sub New()
    MyBase.New()
  End Sub

  'Form overrides dispose to clean up the component list.
  <System.Diagnostics.DebuggerNonUserCode()>
  Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
    If disposing AndAlso components IsNot Nothing Then
      components.Dispose()
    End If
    MyBase.Dispose(disposing)
  End Sub
    Friend WithEvents cmdImpAmazon As NTSButton
    Friend WithEvents LbImpoImpeAmazon As NTSLabel
    Friend WithEvents cmdImpXbio As NTSButton
    Friend WithEvents LbImpoImpeXbio As NTSLabel
    Friend WithEvents LbImpoImpeSleepHome As NTSLabel
    Friend WithEvents cmdImpSleepH As NTSButton
End Class
