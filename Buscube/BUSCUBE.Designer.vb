Partial Public Class FRMBUSCUBE
  Inherits System.Windows.Forms.Form


  'Form overrides dispose to clean up the component list.
  <System.Diagnostics.DebuggerNonUserCode()>
  Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
    Try
      If disposing AndAlso components IsNot Nothing Then
        components.Dispose()
      End If
    Finally
      MyBase.Dispose(disposing)
    End Try
  End Sub

  'Required by the Windows Form Designer
  Private components As System.ComponentModel.IContainer

  <System.Diagnostics.DebuggerStepThrough()>
  Private Sub InitializeComponent()
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FRMBUSCUBE))
    '
    'FRMBUSNET
    '
    components = New System.ComponentModel.Container()
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
    Me.ClientSize = New System.Drawing.Size(124, 32)
    Me.Location = New System.Drawing.Point(-1000, 0)
    Me.Name = "FRMBUSCUBE"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
    Me.Text = "Business Cube"
  End Sub
End Class