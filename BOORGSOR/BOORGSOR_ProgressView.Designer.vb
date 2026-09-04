<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class BOORGSOR_ProgressView
  Inherits FRM__CHIL

  Private components As System.ComponentModel.IContainer

  Friend WithEvents pnProgress As NTSPanel
  Friend WithEvents lbProgress As NTSLabel
  Friend WithEvents pbProgress As NTSProgressBar

  <System.Diagnostics.DebuggerNonUserCode()>
  Public Sub New()
    MyBase.New()
    InitializeComponent()
  End Sub

  <System.Diagnostics.DebuggerNonUserCode()>
  Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
    If disposing AndAlso components IsNot Nothing Then components.Dispose()
    MyBase.Dispose(disposing)
  End Sub

  <System.Diagnostics.DebuggerStepThrough()>
  Public Overridable Sub InitializeComponent()
    Me.pnProgress = New NTSInformatica.NTSPanel()
    Me.lbProgress = New NTSInformatica.NTSLabel()
    Me.pbProgress = New NTSInformatica.NTSProgressBar()
    CType(Me.dttSmartArt, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.pnProgress, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.pnProgress.SuspendLayout()
    CType(Me.pbProgress.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SuspendLayout()
    '
    'pnProgress
    '
    Me.pnProgress.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
    Me.pnProgress.Controls.Add(Me.lbProgress)
    Me.pnProgress.Controls.Add(Me.pbProgress)
    Me.pnProgress.Dock = System.Windows.Forms.DockStyle.Fill
    Me.pnProgress.Location = New System.Drawing.Point(0, 0)
    Me.pnProgress.Name = "pnProgress"
    Me.pnProgress.Size = New System.Drawing.Size(500, 82)
    '
    'lbProgress
    '
    Me.lbProgress.Location = New System.Drawing.Point(10, 10)
    Me.lbProgress.Name = "lbProgress"
    Me.lbProgress.Size = New System.Drawing.Size(480, 35)
    Me.lbProgress.Text = "Avvio importazione..."
    Me.lbProgress.UseMnemonic = False
    '
    'pbProgress
    '
    Me.pbProgress.Location = New System.Drawing.Point(10, 50)
    Me.pbProgress.Name = "pbProgress"
    Me.pbProgress.Properties.Maximum = 100
    Me.pbProgress.Properties.PercentView = True
    Me.pbProgress.Size = New System.Drawing.Size(480, 22)
    '
    'BOORGSOR_ProgressView
    '
    Me.ClientSize = New System.Drawing.Size(500, 82)
    Me.ControlBox = False
    Me.Controls.Add(Me.pnProgress)
    Me.Name = "BOORGSOR_ProgressView"
    Me.Text = "IMPORTAZIONE IMPEGNI IN CORSO..."
    CType(Me.dttSmartArt, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.pnProgress, System.ComponentModel.ISupportInitialize).EndInit()
    Me.pnProgress.ResumeLayout(False)
    CType(Me.pbProgress.Properties, System.ComponentModel.ISupportInitialize).EndInit()
    Me.ResumeLayout(False)
  End Sub
End Class
