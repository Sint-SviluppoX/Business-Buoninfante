Imports System.Data
Imports NTSInformatica.CLN__STD
Imports System
Public Class FROVEBOLL
    Inherits FRMVEBOLL

    Public _oClfBoll As CLFVEBOLL

    Public WithEvents edet_hhtotvol As NTSInformatica.NTSTextBoxNum
    Public WithEvents lbet_hhtotvol As NTSInformatica.NTSLabel
    Public WithEvents ec_hhvol As NTSInformatica.NTSGridColumn
    Public WithEvents edet_hhlprfid As NTSInformatica.NTSTextBoxStr
    Public WithEvents edet_hhtotpn As NTSInformatica.NTSTextBoxNum
    Public WithEvents lbet_hhtotpn As NTSInformatica.NTSLabel
    Public WithEvents ec_hhpn As NTSInformatica.NTSGridColumn
    Public WithEvents edet_hhtotpl As NTSInformatica.NTSTextBoxNum
    Public WithEvents lbet_hhtotpl As NTSInformatica.NTSLabel
    Public WithEvents ec_hhpl As NTSInformatica.NTSGridColumn
    Public WithEvents fm_hhPostiLettoCalcolati As NTSInformatica.NTSGroupBox
    Public WithEvents lb_hhPostiLettoCalcolati As NTSInformatica.NTSLabel
    Public WithEvents ed_hhPostiLettoCalcolati As NTSInformatica.NTSTextBoxNum

    Public Property oClfBoll() As CLFVEBOLL
        Get
            If _oClfBoll Is Nothing Then _oClfBoll = CType(oCleBoll, CLFVEBOLL)
            Return _oClfBoll
        End Get
        Set(ByVal value As CLFVEBOLL)
            _oClfBoll = value
        End Set
    End Property
    Public Overrides Sub InitializeComponent()
        MyBase.InitializeComponent()
        fm_hhPostiLettoCalcolati = New NTSInformatica.NTSGroupBox()
        lb_hhPostiLettoCalcolati = New NTSInformatica.NTSLabel()
        ed_hhPostiLettoCalcolati = New NTSInformatica.NTSTextBoxNum()
        fm_hhPostiLettoCalcolati.Name = "fm_hhPostiLettoCalcolati"
        fm_hhPostiLettoCalcolati.Text = "POSTI LETTO"
        fm_hhPostiLettoCalcolati.Tile = True
        fm_hhPostiLettoCalcolati.TileIndex = 0
        fm_hhPostiLettoCalcolati.Margin = New System.Windows.Forms.Padding(4)
        fm_hhPostiLettoCalcolati.Size = New System.Drawing.Size(200, 80)
        lb_hhPostiLettoCalcolati.Name = "lb_hhPostiLettoCalcolati"
        lb_hhPostiLettoCalcolati.Text = "Posti Letto Calcolati"
        lb_hhPostiLettoCalcolati.BackColor = System.Drawing.Color.Transparent
        lb_hhPostiLettoCalcolati.Location = New System.Drawing.Point(4, 24)
        lb_hhPostiLettoCalcolati.Size = New System.Drawing.Size(188, 20)
        lb_hhPostiLettoCalcolati.NTSBordeStyle = NTSInformatica.NTSLabel.NTSBorderStyle.FieldCaption
        ed_hhPostiLettoCalcolati.Name = "ed_hhPostiLettoCalcolati"
        ed_hhPostiLettoCalcolati.Location = New System.Drawing.Point(4, 48)
        ed_hhPostiLettoCalcolati.Size = New System.Drawing.Size(88, 20)
        ed_hhPostiLettoCalcolati.Properties.AutoHeight = False
        fm_hhPostiLettoCalcolati.Controls.Add(lb_hhPostiLettoCalcolati)
        fm_hhPostiLettoCalcolati.Controls.Add(ed_hhPostiLettoCalcolati)
        flPiede.Controls.Add(fm_hhPostiLettoCalcolati)
        flPiede.Controls.SetChildIndex(fm_hhPostiLettoCalcolati, 0)
    End Sub
    Public Overrides Sub InitControls()
        MyBase.InitControls()
        Try
            ed_hhPostiLettoCalcolati.NTSSetParam(oMenu, "Posti Letto Calcolati", "0.00")
            edet_hhtotvol = CType(NTSFindControlByName(Me, "edet_hhtotvol"), NTSTextBoxNum)
            lbet_hhtotvol = CType(NTSFindControlByName(Me, "lbet_hhtotvol"), NTSLabel)
            edet_hhlprfid = CType(NTSFindControlByName(Me, "edet_hhlprfid"), NTSTextBoxStr)
            ec_hhvol = grvRighe.NTSGetColumnByName("ec_hhvol")
            edet_hhtotvol.NTSSetParam(oMenu, "Volume", oApp.FormatQta)
            ec_hhvol.NTSSetParamNUM(oMenu, "Volume per riga", oApp.FormatQta, 13)
            edet_hhtotpn = CType(NTSFindControlByName(Me, "edet_hhtotpn"), NTSTextBoxNum)
            lbet_hhtotpn = CType(NTSFindControlByName(Me, "lbet_hhtotpn"), NTSLabel)
            ec_hhpn = grvRighe.NTSGetColumnByName("ec_hhpn")
            edet_hhtotpn.NTSSetParam(oMenu, "Peso netto", oApp.FormatQta)
            ec_hhpn.NTSSetParamNUM(oMenu, "Peso netto per riga", oApp.FormatQta, 13)
            edet_hhtotpl = CType(NTSFindControlByName(Me, "edet_hhtotpl"), NTSTextBoxNum)
            lbet_hhtotpl = CType(NTSFindControlByName(Me, "lbet_hhtotpl"), NTSLabel)
            ec_hhpl = grvRighe.NTSGetColumnByName("ec_hhpl")
            edet_hhtotpl.NTSSetParam(oMenu, "Peso lordo", oApp.FormatQta)
            ec_hhpl.NTSSetParamNUM(oMenu, "Peso lordo per riga", oApp.FormatQta, 13)
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
        End Try
    End Sub
    Public Overrides Sub Bindcontrols()
        MyBase.Bindcontrols()
        ed_hhPostiLettoCalcolati.NTSDbField = "TESTA.et_hhPostiLettoCalcolati"
        NTSFormAddDataBinding(dcBoll, fm_hhPostiLettoCalcolati)
    End Sub
    Public Overrides Sub FRMVEBOLL_Load(sender As Object, e As EventArgs)
        MyBase.FRMVEBOLL_Load(sender, e)
        ed_hhPostiLettoCalcolati.Enabled = False
        edet_hhtotvol.Enabled = False
        ec_hhvol.Enabled = False
        lbet_hhtotvol.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        edet_hhlprfid.Enabled = False
        edet_hhtotpn.Enabled = False
        ec_hhpn.Enabled = False
        lbet_hhtotpn.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        edet_hhtotpl.Enabled = False
        ec_hhpl.Enabled = False
        lbet_hhtotpl.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    End Sub

End Class
