Imports System.Data
Imports System.IO
Imports NTSInformatica.CLN__STD
Imports System
Public Class FROORGSOR
  Inherits FRMORGSOR
  Public _oClfGsor As CLFORGSOR
  Public WithEvents edet_hhtotvol As NTSInformatica.NTSTextBoxNum
  Public WithEvents lbet_hhtotvol As NTSInformatica.NTSLabel
  Public WithEvents ec_hhvol As NTSInformatica.NTSGridColumn
  Public WithEvents ec_hhisetrfid As NTSInformatica.NTSGridColumn
  Public WithEvents edet_hhtotpn As NTSInformatica.NTSTextBoxNum
  Public WithEvents lbet_hhtotpn As NTSInformatica.NTSLabel
  Public WithEvents ec_hhpn As NTSInformatica.NTSGridColumn
  Public WithEvents edet_hhtotpl As NTSInformatica.NTSTextBoxNum
  Public WithEvents lbet_hhtotpl As NTSInformatica.NTSLabel
  Public WithEvents ec_hhpl As NTSInformatica.NTSGridColumn
  Public WithEvents tlb_hhImportaImpegni As NTSInformatica.NTSBarButtonItem
  Private Shared strUltimaCartellaImport As String = ""
  Public Property oClfGsor() As CLFORGSOR
    Get
      If _oClfGsor Is Nothing Then _oClfGsor = CType(oCleGsor, CLFORGSOR)
      Return _oClfGsor
    End Get
    Set(ByVal value As CLFORGSOR)
      _oClfGsor = value
    End Set
  End Property
  Public Overrides Sub InitializeComponent()
    MyBase.InitializeComponent()
    Me.tlb_hhImportaImpegni = New NTSInformatica.NTSBarButtonItem()
    Me.NtsBarManager1.Items.AddRange(New DevExpress.XtraBars.BarItem() {Me.tlb_hhImportaImpegni})
    Me.NtsBarManager1.MaxItemId = 214
    Me.tlbMain.LinksPersistInfo.AddRange(New DevExpress.XtraBars.LinkPersistInfo() {
      New DevExpress.XtraBars.LinkPersistInfo(Me.tlb_hhImportaImpegni)})
    '
    'tlb_hhImportaImpegni
    '
    Me.tlb_hhImportaImpegni.Caption = "Importa impegni"
    Me.tlb_hhImportaImpegni.Glyph = GetGlyphImportaImpegni()
    Me.tlb_hhImportaImpegni.Id = 213
    Me.tlb_hhImportaImpegni.Name = "tlb_hhImportaImpegni"
    Me.tlb_hhImportaImpegni.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionInMenu
    Me.tlb_hhImportaImpegni.Visible = False
  End Sub
  Public Overrides Sub InitControls()
    MyBase.InitControls()
    Try
      edet_hhtotvol = CType(NTSFindControlByName(Me, "edet_hhtotvol"), NTSTextBoxNum)
      lbet_hhtotvol = CType(NTSFindControlByName(Me, "lbet_hhtotvol"), NTSLabel)
      ec_hhvol = grvRighe.NTSGetColumnByName("ec_hhvol")
      edet_hhtotvol.NTSSetParam(oMenu, "Volume", oApp.FormatQta)
      ec_hhvol.NTSSetParamNUM(oMenu, "Volume per riga", oApp.FormatQta, 13)
      ec_hhisetrfid = grvRighe.NTSGetColumnByName("ec_hhisetrfid")
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
      '-------------------------------------------------
      CLN__STD.GestErr(ex, Me, "")
      '-------------------------------------------------
    End Try
  End Sub
  Public Overrides Sub FRMORGSOR_Load(sender As Object, e As EventArgs)
    Try
      MyBase.FRMORGSOR_Load(sender, e)
      edet_hhtotvol.Enabled = False
      ec_hhvol.Enabled = False
      'ec_hhisetrfid.Visible = CBool(IIf(cbTipoDoc.SelectedValue.ToString <> "H", False, True))
      lbet_hhtotvol.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      edet_hhtotpn.Enabled = False
      ec_hhpn.Enabled = False
      lbet_hhtotpn.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      edet_hhtotpl.Enabled = False
      ec_hhpl.Enabled = False
      lbet_hhtotpl.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Catch ex As Exception
      CLN__STD.GestErr(ex, Me, "")
    End Try
  End Sub
  Public Overrides Sub cbTipoDoc_SelectedIndexChanged(sender As Object, e As EventArgs)
    Try
      MyBase.cbTipoDoc_SelectedIndexChanged(sender, e)
      'Se Ordine di Produzione abilito, in griglia, il campo/colonna pers. che indica se � una riga che dovr� andare nelle etichette rfid
      If Not IsNothing(ec_hhisetrfid) Then
        If cbTipoDoc.SelectedValue.ToString <> "H" Then
          ec_hhisetrfid.Enabled = False
        Else
          ec_hhisetrfid.Enabled = True
        End If
      End If
      AggiornaPulsanteImporta()
    Catch ex As Exception
      '-------------------------------------------------
      CLN__STD.GestErr(ex, Me, "")
    '-------------------------------------------------	
    End Try
  End Sub
#Region "Importazione impegni"

  Private Sub AggiornaPulsanteImportaDaNumero(sender As Object, e As EventArgs) Handles edNumDoc.Validated
    Try
      AggiornaPulsanteImporta()
    Catch ex As Exception
      CLN__STD.GestErr(ex, Me, "")
    End Try
  End Sub
  Public Overridable Sub tlb_hhImportaImpegni_ItemClick(ByVal sender As Object,
    ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles tlb_hhImportaImpegni.ItemClick
    Dim frmProgress As BOORGSOR_ProgressView = Nothing
    Dim importatore As HH_ImportatoreImpegni = Nothing
    Dim esito As HH_EsitoImportazioneImpegni = Nothing
    Try
      Dim percorsoFile As String = SelezionaFileImpegni()
      If String.IsNullOrWhiteSpace(percorsoFile) Then Return

      Me.tlb_hhImportaImpegni.Enabled = False
      Try
        frmProgress = New BOORGSOR_ProgressView()
        If Not frmProgress.Init(CType(oApp.oMenu, CLE__MENU), Nothing, oApp.Ditta) Then Return

        importatore = New HH_ImportatoreImpegni(oApp, CType(oApp.oMenu, CLE__MENU), oClfGsor)
        AddHandler importatore.Avanzamento, AddressOf frmProgress.Aggiorna
        frmProgress.Show(Me)
        frmProgress.Aggiorna("Avvio importazione impegni...", 0)

        esito = importatore.Importa(percorsoFile)
      Finally
        If importatore IsNot Nothing AndAlso frmProgress IsNot Nothing Then
          RemoveHandler importatore.Avanzamento, AddressOf frmProgress.Aggiorna
        End If
        If frmProgress IsNot Nothing Then frmProgress.Close()
        Me.tlb_hhImportaImpegni.Enabled = True
      End Try

      MostraEsitoImportazione(esito)
    Catch ex As Exception
      CLN__STD.GestErr(ex, Me, "")
    End Try
  End Sub

#End Region

#Region "Utils"

  Private Sub AggiornaPulsanteImporta()
    Try
      If tlb_hhImportaImpegni Is Nothing Then Return

      Dim strTipoDoc As String = ""
      If oCleGsor IsNot Nothing AndAlso oCleGsor.dttET IsNot Nothing AndAlso
         oCleGsor.dttET.Rows.Count > 0 Then
        strTipoDoc = NTSCStr(oCleGsor.dttET.Rows(0)!et_tipork)
      ElseIf cbTipoDoc IsNot Nothing Then
        strTipoDoc = NTSCStr(cbTipoDoc.SelectedValue)
      End If

      Me.tlb_hhImportaImpegni.Visible = String.Equals(strTipoDoc.Trim(), "R",
                                                       StringComparison.OrdinalIgnoreCase)
    Catch ex As Exception
      CLN__STD.GestErr(ex, Me, "")
    End Try
  End Sub

  Private Function SelezionaFileImpegni() As String
    Try
      Using dialogo As New OpenFileDialog()
        dialogo.Title = "Seleziona il file degli impegni clienti"
        dialogo.Filter = "File Excel 97-2003 (*.xls)|*.xls|Tutti i file (*.*)|*.*"
        dialogo.CheckFileExists = True
        dialogo.Multiselect = False
        If Not String.IsNullOrWhiteSpace(strUltimaCartellaImport) AndAlso
           Directory.Exists(strUltimaCartellaImport) Then
          dialogo.InitialDirectory = strUltimaCartellaImport
        End If

        If dialogo.ShowDialog(Me) <> DialogResult.OK Then Return ""
        strUltimaCartellaImport = Path.GetDirectoryName(dialogo.FileName)
        Return dialogo.FileName
      End Using
    Catch ex As Exception
      CLN__STD.GestErr(ex, Me, "")
      Return ""
    End Try
  End Function

  Private Sub MostraEsitoImportazione(ByVal esito As HH_EsitoImportazioneImpegni)
    Try
      If esito Is Nothing Then
        oApp.MsgBoxErr("Importazione impegni non completata.")
        Return
      End If

      If esito.ErroriDocumento.Count = 0 Then
        oApp.MsgBoxInfo("Importazione completata. Impegni salvati: " &
                        esito.DocumentiSalvati.ToString() & ".")
        Return
      End If

      Dim messaggio As String = "Impegni salvati: " & esito.DocumentiSalvati.ToString() & "." &
                                vbCrLf & "Documenti non salvati:" & vbCrLf &
                                String.Join(vbCrLf, esito.ErroriDocumento.ToArray())
      oApp.MsgBoxExclamation(messaggio)
    Catch ex As Exception
      CLN__STD.GestErr(ex, Me, "")
    End Try
  End Sub

  Private Function GetGlyphImportaImpegni() As Image
    Try
      Using stream As Stream = GetType(FROORGSOR).Assembly.GetManifestResourceStream(
        "NTSInformatica.ImportaImpegni24.png")
        If stream Is Nothing Then Return Nothing
        Using img As Image = Image.FromStream(stream)
          Return New Bitmap(img)
        End Using
      End Using
    Catch ex As Exception
      CLN__STD.GestErr(ex, Me, "")
      Return Nothing
    End Try
  End Function

#End Region
  'Public Overrides Function Salva(nStampa As Integer, bAllowCancel As Boolean) As Boolean
  '  Try
  '    'Lancio prima il calcolo dei totali
  '    oCleGsor.dttET.Rows(0)!et_hhtotvol = oClfGsor.CalcolaTotaleVolume()
  '    Return MyBase.Salva(nStampa, bAllowCancel)
  '  Catch ex As Exception
  '    Dim strErr As String = GestError(ex, Me, "", oApp.InfoError, oApp.ErrorLogFile, True)
  '  End Try
  'End Function


End Class
