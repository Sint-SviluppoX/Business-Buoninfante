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
  Public Property oClfBoll() As CLFVEBOLL
    Get
      If _oClfBoll Is Nothing Then _oClfBoll = CType(ocleBoll, CLFVEBOLL)
      Return _oClfBoll
    End Get
    Set(ByVal value As CLFVEBOLL)
      _oClfBoll = value
    End Set
  End Property
  Public Overrides Sub InitControls()
    MyBase.InitControls()
    Try
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
  Public Overrides Sub FRMVEBOLL_Load(sender As Object, e As EventArgs)
    MyBase.FRMVEBOLL_Load(sender, e)
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