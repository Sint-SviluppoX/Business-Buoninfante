Imports System.Data
Imports NTSInformatica.CLN__STD

Public Class FRO__DESG
  Inherits FRM__DESG

#Region "Controlli importazione impegni"

    Public WithEvents fm_hhImpExc As NTSGroupBox
    Public WithEvents lb_hhCodDestExc As NTSLabel
    Public WithEvents ed_hhCodDestExc As NTSTextBoxNum
    Public WithEvents lb_hhGiornoConsegna As NTSLabel
    Public WithEvents cb_hhGiornoConsegna As NTSComboBox

    Public Overrides Sub InitializeComponent()
        Try
            MyBase.InitializeComponent()

            Me.fm_hhImpExc = New NTSInformatica.NTSGroupBox()
            Me.lb_hhCodDestExc = New NTSInformatica.NTSLabel()
            Me.ed_hhCodDestExc = New NTSInformatica.NTSTextBoxNum()
            Me.lb_hhGiornoConsegna = New NTSInformatica.NTSLabel()
            Me.cb_hhGiornoConsegna = New NTSInformatica.NTSComboBox()
            CType(Me.fm_hhImpExc, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.fm_hhImpExc.SuspendLayout()
            CType(Me.ed_hhCodDestExc.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.cb_hhGiornoConsegna.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.flAltriDati.SuspendLayout()
            Me.SuspendLayout()
            '
            'fm_hhImpExc
            '
            Me.fm_hhImpExc.AllowDrop = True
            Me.fm_hhImpExc.Appearance.BackColor = System.Drawing.Color.Transparent
            Me.fm_hhImpExc.Appearance.Options.UseBackColor = True
            Me.fm_hhImpExc.Controls.Add(Me.lb_hhCodDestExc)
            Me.fm_hhImpExc.Controls.Add(Me.ed_hhCodDestExc)
            Me.fm_hhImpExc.Controls.Add(Me.lb_hhGiornoConsegna)
            Me.fm_hhImpExc.Controls.Add(Me.cb_hhGiornoConsegna)
            Me.fm_hhImpExc.Location = New System.Drawing.Point(4, 516)
            Me.fm_hhImpExc.Margin = New System.Windows.Forms.Padding(4)
            Me.fm_hhImpExc.Name = "fm_hhImpExc"
            Me.fm_hhImpExc.Size = New System.Drawing.Size(456, 120)
            Me.fm_hhImpExc.Text = "IMPORT EXCEL"
            Me.fm_hhImpExc.Tile = True
            Me.fm_hhImpExc.TileIndex = 5
            '
            'lb_hhCodDestExc
            '
            Me.lb_hhCodDestExc.BackColor = System.Drawing.Color.Transparent
            Me.lb_hhCodDestExc.Location = New System.Drawing.Point(4, 28)
            Me.lb_hhCodDestExc.Name = "lb_hhCodDestExc"
            Me.lb_hhCodDestExc.NTSBordeStyle = NTSInformatica.NTSLabel.NTSBorderStyle.FieldCaption
            Me.lb_hhCodDestExc.Size = New System.Drawing.Size(132, 20)
            Me.lb_hhCodDestExc.Text = "Codice destinazione excel"
            Me.lb_hhCodDestExc.UseMnemonic = False
            '
            'ed_hhCodDestExc
            '
            Me.ed_hhCodDestExc.Location = New System.Drawing.Point(144, 28)
            Me.ed_hhCodDestExc.Name = "ed_hhCodDestExc"
            Me.ed_hhCodDestExc.Properties.AutoHeight = False
            Me.ed_hhCodDestExc.Size = New System.Drawing.Size(40, 20)
            '
            'lb_hhGiornoConsegna
            '
            Me.lb_hhGiornoConsegna.BackColor = System.Drawing.Color.Transparent
            Me.lb_hhGiornoConsegna.Location = New System.Drawing.Point(4, 56)
            Me.lb_hhGiornoConsegna.Name = "lb_hhGiornoConsegna"
            Me.lb_hhGiornoConsegna.NTSBordeStyle = NTSInformatica.NTSLabel.NTSBorderStyle.FieldCaption
            Me.lb_hhGiornoConsegna.Size = New System.Drawing.Size(88, 20)
            Me.lb_hhGiornoConsegna.Text = "Giorno consegna"
            Me.lb_hhGiornoConsegna.UseMnemonic = False
            '
            'cb_hhGiornoConsegna
            '
            Me.cb_hhGiornoConsegna.Location = New System.Drawing.Point(100, 56)
            Me.cb_hhGiornoConsegna.Name = "cb_hhGiornoConsegna"
            Me.cb_hhGiornoConsegna.Properties.AutoHeight = False
            Me.cb_hhGiornoConsegna.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {
        New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
            Me.cb_hhGiornoConsegna.Properties.DropDownRows = 8
            Me.cb_hhGiornoConsegna.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
            Me.cb_hhGiornoConsegna.Size = New System.Drawing.Size(100, 20)
            '
            'FRO__DESG
            '
            Me.flAltriDati.Controls.Add(Me.fm_hhImpExc)
            CType(Me.fm_hhImpExc, System.ComponentModel.ISupportInitialize).EndInit()
            Me.fm_hhImpExc.ResumeLayout(False)
            CType(Me.ed_hhCodDestExc.Properties, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.cb_hhGiornoConsegna.Properties, System.ComponentModel.ISupportInitialize).EndInit()
            Me.flAltriDati.ResumeLayout(False)
            Me.ResumeLayout(False)
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
        End Try
    End Sub

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

#End Region

End Class
