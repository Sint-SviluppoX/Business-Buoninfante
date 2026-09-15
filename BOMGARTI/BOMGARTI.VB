Imports System.Data
Imports NTSInformatica.CLN__STD
Imports System.Globalization

Imports System
Public Class FROMGARTI
  Inherits FRMMGARTI

  Public WithEvents fm_hhPostiLetto As NTSGroupBox
  Public WithEvents lb_hhPostiLetto As NTSLabel
  Public WithEvents ed_hhPostiLetto As NTSTextBoxNum

  Public Overrides Sub InitializeComponent()
    MyBase.InitializeComponent()

    fm_hhPostiLetto = New NTSGroupBox()
    lb_hhPostiLetto = New NTSLabel()
    ed_hhPostiLetto = New NTSTextBoxNum()

    fm_hhPostiLetto.Name = "fm_hhPostiLetto"
    fm_hhPostiLetto.Text = "POSTI LETTO"
    fm_hhPostiLetto.Tile = True
    fm_hhPostiLetto.TileIndex = 5
    fm_hhPostiLetto.Margin = New System.Windows.Forms.Padding(4)
    fm_hhPostiLetto.Size = New System.Drawing.Size(456, 72)

    lb_hhPostiLetto.Name = "lb_hhPostiLetto"
    lb_hhPostiLetto.Text = "Posti Letto"
    lb_hhPostiLetto.BackColor = System.Drawing.Color.Transparent
    lb_hhPostiLetto.Location = New System.Drawing.Point(4, 28)
    lb_hhPostiLetto.Size = New System.Drawing.Size(100, 20)
    lb_hhPostiLetto.NTSBordeStyle = NTSLabel.NTSBorderStyle.FieldCaption

    ed_hhPostiLetto.Name = "ed_hhPostiLetto"
    ed_hhPostiLetto.Location = New System.Drawing.Point(112, 28)
    ed_hhPostiLetto.Size = New System.Drawing.Size(88, 20)
    ed_hhPostiLetto.Properties.AutoHeight = False

    fm_hhPostiLetto.Controls.Add(lb_hhPostiLetto)
    fm_hhPostiLetto.Controls.Add(ed_hhPostiLetto)
    flPage6.Controls.Add(fm_hhPostiLetto)
  End Sub

  Public Overrides Sub InitControls()
    MyBase.InitControls()
    'L'opzione globale usa 5 come limite iniziale; i valori configurati prevalgono.
    Dim massimo As Decimal = 5D
    Dim valoreOpzione As String = oClfArti.oClhArti.GetSettingBus("BSMGARTI", "OPZIONI", ".",
                                                 "LimiteMassimoPostiLetto", "", "",
                                                 "")
    If String.IsNullOrWhiteSpace(valoreOpzione) Then
      valoreOpzione = "5"
      oClfArti.oClhArti.SaveSettingBus("BSMGARTI", "OPZIONI", ".",
                                      "LimiteMassimoPostiLetto", valoreOpzione, "",
                                      False, False, False)
    End If
    Dim massimoConfigurato As Decimal
    Dim stileNumero As NumberStyles = NumberStyles.AllowDecimalPoint Or NumberStyles.AllowLeadingSign
    If Decimal.TryParse(valoreOpzione, stileNumero, CultureInfo.InvariantCulture,
                        massimoConfigurato) OrElse
       Decimal.TryParse(valoreOpzione, stileNumero, CultureInfo.CurrentCulture,
                        massimoConfigurato) Then
      If massimoConfigurato >= 0D AndAlso massimoConfigurato <= 9999999999999999.99D Then
        massimo = massimoConfigurato
      End If
    End If
    ed_hhPostiLetto.NTSSetParam(oMenu, "Posti Letto", "0.00", 18, 0D, massimo)
  End Sub

  Public Overrides Sub Bindcontrols()
    MyBase.Bindcontrols()
    ed_hhPostiLetto.NTSDbField = "ARTICO.ar_hhPostiLetto"
    NTSFormAddDataBinding(dcArti, fm_hhPostiLetto)
  End Sub

  Public _oClfArti As CLFMGARTI
  Public Property oClfArti() As CLFMGARTI
    Get
      If _oClfArti Is Nothing Then _oClfArti = CType(ocleArti, CLFMGARTI)
      Return _oClfArti
    End Get
    Set(ByVal value As CLFMGARTI)
      _oClfArti = value
    End Set
  End Property

End Class

