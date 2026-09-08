Imports System.Data
Imports NTSInformatica.CLN__STD
Imports NTSInformatica.CLD__BASE
Imports System
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Net
Imports System.Collections
Imports System.Collections.Specialized
Imports System.Text
Imports System.IO
Imports System.Reflection

Public Class FRM__HHID

#Region "Controlli"
    'NtsBar
    Public WithEvents NtsBarManager1 As NTSInformatica.NTSBarManager
    Public WithEvents barDockControlTop As DevExpress.XtraBars.BarDockControl
    Public WithEvents barDockControlBottom As DevExpress.XtraBars.BarDockControl
    Public WithEvents barDockControlLeft As DevExpress.XtraBars.BarDockControl
    Public WithEvents barDockControlRight As DevExpress.XtraBars.BarDockControl
    Public WithEvents NtsBar1 As NTSInformatica.NTSBar
    Public WithEvents tlbZoom As NTSInformatica.NTSBarButtonItem
    Public WithEvents tlbEsci As NTSInformatica.NTSBarButtonItem
    Public WithEvents tlbLog As NTSInformatica.NTSBarButtonItem
    Public WithEvents fmTipoDato As NTSInformatica.NTSGroupBox

    'NtsLabel
    Public WithEvents lbProgress As NTSInformatica.NTSLabel

    'NtsPanel
    Public WithEvents pnPrincipale As NTSInformatica.NTSPanel
    Public WithEvents pnProgress As NTSInformatica.NTSPanel

    'Altro
    Public WithEvents pbProgress As NTSInformatica.NTSProgressBar
    Public WithEvents OFD_File As NTSInformatica.NTSOpenFileDialog
    Private components As System.ComponentModel.IContainer

#End Region

#Region "Varianti"
    '*** Base ***
    Public oCallParams As CLE__CLDP
    Public oCleHhId As New CLE__HHID
    Public dsHhId As New DataSet

    '*** Thread ***
    Public ThreadCreaDoc As System.Threading.Thread
    Public ThreadImportaArticoli As System.Threading.Thread
    Public ThreadImportaListini As System.Threading.Thread

    '*** Altro ***
    Public nCodListino As String = String.Empty
    Public DataListino As String = String.Empty

    Public nCreati As Integer = 0
    Public nDuplicati As Integer = 0
    Public nErrori As Integer = 0

    Public bCreaArticoNonEsistente As Boolean = False
#End Region

    Public Overloads Function Init(ByRef Menu As CLE__MENU, ByRef Param As CLE__CLDP, Optional ByVal Ditta As String = "", Optional ByRef SharedControls As CLE__EVNT = Nothing) As Boolean
        oMenu = Menu
        oApp = oMenu.App
        oCallParams = Param
        If Ditta <> "" Then
            DittaCorrente = Ditta
        Else
            DittaCorrente = oApp.Ditta
        End If
        Me.GctlTipoDoc = ""
        InitializeComponent()
        Me.MinimumSize = Me.Size

        '*** Versione del programma ***
        Dim FullVersion As String() = Assembly.GetExecutingAssembly().GetName().Version.ToString().Split("."c)
        Dim GCVersion As String = FullVersion(1) & "." & FullVersion(2) & "." & FullVersion(3)
        Me.Text &= " v." & GCVersion

        '------------------------------------------------
        'creo e attivo l'entity e inizializzo la funzione che dovrà rilevare gli eventi dall'ENTITY
        Dim strErr As String = ""
        Dim oTmp As Object = Nothing
        If CLN__STD.NTSIstanziaDll(oApp.ServerDir, oApp.NetDir, "BN__HHID", "BE__HHID", oTmp, strErr, False, "", "") = False Then
            oApp.MsgBoxErr(oApp.Tr(Me, 128271029889882656, "ERRORE in fase di creazione Entity:" & vbCrLf & "|" & strErr & "|"))
            Return False
        End If
        oCleHhId = CType(oTmp, CLE__HHID)
        '------------------------------------------------
        AddHandler oCleHhId.RemoteEvent, AddressOf GestisciEventiEntity
        If oCleHhId.Init(oApp, NTSScript, oMenu.oCleComm, "", False, "", "") = False Then Return False

        Return True
    End Function
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FRM__HHID))
        Me.NtsBarManager1 = New NTSInformatica.NTSBarManager()
        Me.NtsBar1 = New NTSInformatica.NTSBar()
        Me.tlbLog = New NTSInformatica.NTSBarButtonItem()
        Me.tlbZoom = New NTSInformatica.NTSBarButtonItem()
        Me.tlbEsci = New NTSInformatica.NTSBarButtonItem()
        Me.barDockControlTop = New DevExpress.XtraBars.BarDockControl()
        Me.barDockControlBottom = New DevExpress.XtraBars.BarDockControl()
        Me.barDockControlLeft = New DevExpress.XtraBars.BarDockControl()
        Me.barDockControlRight = New DevExpress.XtraBars.BarDockControl()
        Me.fmTipoDato = New NTSInformatica.NTSGroupBox()
        Me.LbImpoImpeSleepHome = New NTSInformatica.NTSLabel()
        Me.cmdImpSleepH = New NTSInformatica.NTSButton()
        Me.cmdImpXbio = New NTSInformatica.NTSButton()
        Me.LbImpoImpeXbio = New NTSInformatica.NTSLabel()
        Me.LbImpoImpeAmazon = New NTSInformatica.NTSLabel()
        Me.cmdImpAmazon = New NTSInformatica.NTSButton()
        Me.pnProgress = New NTSInformatica.NTSPanel()
        Me.lbProgress = New NTSInformatica.NTSLabel()
        Me.pbProgress = New NTSInformatica.NTSProgressBar()
        Me.pnPrincipale = New NTSInformatica.NTSPanel()
        CType(Me.dttSmartArt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NtsBarManager1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.fmTipoDato, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fmTipoDato.SuspendLayout()
        CType(Me.pnProgress, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnProgress.SuspendLayout()
        CType(Me.pbProgress.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pnPrincipale, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnPrincipale.SuspendLayout()
        Me.SuspendLayout()
        '
        'NtsBarManager1
        '
        Me.NtsBarManager1.AllowCustomization = False
        Me.NtsBarManager1.Bars.AddRange(New DevExpress.XtraBars.Bar() {Me.NtsBar1})
        Me.NtsBarManager1.DockControls.Add(Me.barDockControlTop)
        Me.NtsBarManager1.DockControls.Add(Me.barDockControlBottom)
        Me.NtsBarManager1.DockControls.Add(Me.barDockControlLeft)
        Me.NtsBarManager1.DockControls.Add(Me.barDockControlRight)
        Me.NtsBarManager1.Form = Me
        Me.NtsBarManager1.Items.AddRange(New DevExpress.XtraBars.BarItem() {Me.tlbZoom, Me.tlbEsci, Me.tlbLog})
        Me.NtsBarManager1.MaxItemId = 17
        '
        'NtsBar1
        '
        Me.NtsBar1.BarName = "tlbMain"
        Me.NtsBar1.DockCol = 0
        Me.NtsBar1.DockRow = 0
        Me.NtsBar1.DockStyle = DevExpress.XtraBars.BarDockStyle.Top
        Me.NtsBar1.LinksPersistInfo.AddRange(New DevExpress.XtraBars.LinkPersistInfo() {New DevExpress.XtraBars.LinkPersistInfo(Me.tlbLog), New DevExpress.XtraBars.LinkPersistInfo(Me.tlbZoom), New DevExpress.XtraBars.LinkPersistInfo(Me.tlbEsci)})
        Me.NtsBar1.OptionsBar.AllowQuickCustomization = False
        Me.NtsBar1.OptionsBar.DisableClose = True
        Me.NtsBar1.OptionsBar.DrawDragBorder = False
        Me.NtsBar1.OptionsBar.UseWholeRow = True
        Me.NtsBar1.Text = "tlbMain"
        '
        'tlbLog
        '
        Me.tlbLog.Caption = "Log"
        Me.tlbLog.Glyph = CType(resources.GetObject("tlbLog.Glyph"), System.Drawing.Image)
        Me.tlbLog.Id = 16
        Me.tlbLog.Name = "tlbLog"
        Me.tlbLog.Visible = True
        '
        'tlbZoom
        '
        Me.tlbZoom.Caption = "Zoom"
        Me.tlbZoom.Glyph = CType(resources.GetObject("tlbZoom.Glyph"), System.Drawing.Image)
        Me.tlbZoom.Id = 6
        Me.tlbZoom.Name = "tlbZoom"
        Me.tlbZoom.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionInMenu
        Me.tlbZoom.Visible = True
        '
        'tlbEsci
        '
        Me.tlbEsci.Caption = "Esci"
        Me.tlbEsci.Glyph = CType(resources.GetObject("tlbEsci.Glyph"), System.Drawing.Image)
        Me.tlbEsci.Id = 11
        Me.tlbEsci.Name = "tlbEsci"
        Me.tlbEsci.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionInMenu
        Me.tlbEsci.Visible = True
        '
        'barDockControlTop
        '
        Me.barDockControlTop.CausesValidation = False
        Me.barDockControlTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.barDockControlTop.Location = New System.Drawing.Point(0, 0)
        Me.barDockControlTop.Size = New System.Drawing.Size(361, 35)
        '
        'barDockControlBottom
        '
        Me.barDockControlBottom.CausesValidation = False
        Me.barDockControlBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.barDockControlBottom.Location = New System.Drawing.Point(0, 255)
        Me.barDockControlBottom.Size = New System.Drawing.Size(361, 0)
        '
        'barDockControlLeft
        '
        Me.barDockControlLeft.CausesValidation = False
        Me.barDockControlLeft.Dock = System.Windows.Forms.DockStyle.Left
        Me.barDockControlLeft.Location = New System.Drawing.Point(0, 35)
        Me.barDockControlLeft.Size = New System.Drawing.Size(0, 220)
        '
        'barDockControlRight
        '
        Me.barDockControlRight.CausesValidation = False
        Me.barDockControlRight.Dock = System.Windows.Forms.DockStyle.Right
        Me.barDockControlRight.Location = New System.Drawing.Point(361, 35)
        Me.barDockControlRight.Size = New System.Drawing.Size(0, 220)
        '
        'fmTipoDato
        '
        Me.fmTipoDato.AllowDrop = True
        Me.fmTipoDato.Appearance.BackColor = System.Drawing.Color.Transparent
        Me.fmTipoDato.Appearance.Options.UseBackColor = True
        Me.fmTipoDato.Controls.Add(Me.LbImpoImpeSleepHome)
        Me.fmTipoDato.Controls.Add(Me.cmdImpSleepH)
        Me.fmTipoDato.Controls.Add(Me.cmdImpXbio)
        Me.fmTipoDato.Controls.Add(Me.LbImpoImpeXbio)
        Me.fmTipoDato.Controls.Add(Me.LbImpoImpeAmazon)
        Me.fmTipoDato.Controls.Add(Me.cmdImpAmazon)
        Me.fmTipoDato.Location = New System.Drawing.Point(4, 8)
        Me.fmTipoDato.Name = "fmTipoDato"
        Me.fmTipoDato.Size = New System.Drawing.Size(352, 152)
        Me.fmTipoDato.Text = "IMPORTAZIONE DATI DA  WEB"
        '
        'LbImpoImpeSleepHome
        '
        Me.LbImpoImpeSleepHome.BackColor = System.Drawing.Color.Transparent
        Me.LbImpoImpeSleepHome.Location = New System.Drawing.Point(236, 24)
        Me.LbImpoImpeSleepHome.Name = "LbImpoImpeSleepHome"
        Me.LbImpoImpeSleepHome.NTSBordeStyle = NTSInformatica.NTSLabel.NTSBorderStyle.FieldCaption
        Me.LbImpoImpeSleepHome.Size = New System.Drawing.Size(112, 20)
        Me.LbImpoImpeSleepHome.Text = "Sleep Home"
        Me.LbImpoImpeSleepHome.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.LbImpoImpeSleepHome.UseMnemonic = False
        '
        'cmdImpSleepH
        '
        Me.cmdImpSleepH.Image = CType(resources.GetObject("cmdImpSleepH.Image"), System.Drawing.Image)
        Me.cmdImpSleepH.ImageLocation = DevExpress.XtraEditors.ImageLocation.MiddleCenter
        Me.cmdImpSleepH.Location = New System.Drawing.Point(236, 48)
        Me.cmdImpSleepH.Name = "cmdImpSleepH"
        Me.cmdImpSleepH.Size = New System.Drawing.Size(112, 100)
        Me.cmdImpSleepH.Text = "Impo Xbio/S.Home"
        '
        'cmdImpXbio
        '
        Me.cmdImpXbio.Image = CType(resources.GetObject("cmdImpXbio.Image"), System.Drawing.Image)
        Me.cmdImpXbio.ImageLocation = DevExpress.XtraEditors.ImageLocation.MiddleCenter
        Me.cmdImpXbio.Location = New System.Drawing.Point(120, 48)
        Me.cmdImpXbio.Name = "cmdImpXbio"
        Me.cmdImpXbio.Size = New System.Drawing.Size(112, 100)
        Me.cmdImpXbio.Text = "Impo Xbio/S.Home"
        '
        'LbImpoImpeXbio
        '
        Me.LbImpoImpeXbio.BackColor = System.Drawing.Color.Transparent
        Me.LbImpoImpeXbio.Location = New System.Drawing.Point(120, 24)
        Me.LbImpoImpeXbio.Name = "LbImpoImpeXbio"
        Me.LbImpoImpeXbio.NTSBordeStyle = NTSInformatica.NTSLabel.NTSBorderStyle.FieldCaption
        Me.LbImpoImpeXbio.Size = New System.Drawing.Size(112, 20)
        Me.LbImpoImpeXbio.Text = "Xbio"
        Me.LbImpoImpeXbio.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.LbImpoImpeXbio.UseMnemonic = False
        '
        'LbImpoImpeAmazon
        '
        Me.LbImpoImpeAmazon.BackColor = System.Drawing.Color.Transparent
        Me.LbImpoImpeAmazon.Location = New System.Drawing.Point(4, 24)
        Me.LbImpoImpeAmazon.Name = "LbImpoImpeAmazon"
        Me.LbImpoImpeAmazon.NTSBordeStyle = NTSInformatica.NTSLabel.NTSBorderStyle.FieldCaption
        Me.LbImpoImpeAmazon.Size = New System.Drawing.Size(112, 20)
        Me.LbImpoImpeAmazon.Text = "Amazon"
        Me.LbImpoImpeAmazon.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.LbImpoImpeAmazon.UseMnemonic = False
        '
        'cmdImpAmazon
        '
        Me.cmdImpAmazon.Image = CType(resources.GetObject("cmdImpAmazon.Image"), System.Drawing.Image)
        Me.cmdImpAmazon.ImageLocation = DevExpress.XtraEditors.ImageLocation.MiddleCenter
        Me.cmdImpAmazon.Location = New System.Drawing.Point(4, 48)
        Me.cmdImpAmazon.Name = "cmdImpAmazon"
        Me.cmdImpAmazon.Size = New System.Drawing.Size(112, 100)
        Me.cmdImpAmazon.Text = "Impo. Amazon"
        '
        'pnProgress
        '
        Me.pnProgress.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.pnProgress.Controls.Add(Me.lbProgress)
        Me.pnProgress.Controls.Add(Me.pbProgress)
        Me.pnProgress.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnProgress.Location = New System.Drawing.Point(0, 203)
        Me.pnProgress.Name = "pnProgress"
        Me.pnProgress.Size = New System.Drawing.Size(361, 52)
        '
        'lbProgress
        '
        Me.lbProgress.BackColor = System.Drawing.Color.Transparent
        Me.lbProgress.Location = New System.Drawing.Point(4, 4)
        Me.lbProgress.Name = "lbProgress"
        Me.lbProgress.Size = New System.Drawing.Size(352, 20)
        Me.lbProgress.Text = "0/0"
        Me.lbProgress.UseMnemonic = False
        '
        'pbProgress
        '
        Me.pbProgress.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pbProgress.Location = New System.Drawing.Point(4, 28)
        Me.pbProgress.Name = "pbProgress"
        Me.pbProgress.Size = New System.Drawing.Size(353, 20)
        Me.pbProgress.TabIndex = 0
        '
        'pnPrincipale
        '
        Me.pnPrincipale.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.pnPrincipale.Controls.Add(Me.fmTipoDato)
        Me.pnPrincipale.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnPrincipale.Location = New System.Drawing.Point(0, 35)
        Me.pnPrincipale.Name = "pnPrincipale"
        Me.pnPrincipale.Size = New System.Drawing.Size(361, 168)
        '
        'FRM__HHID
        '
        Me.ClientSize = New System.Drawing.Size(361, 255)
        Me.Controls.Add(Me.pnPrincipale)
        Me.Controls.Add(Me.pnProgress)
        Me.Controls.Add(Me.barDockControlLeft)
        Me.Controls.Add(Me.barDockControlRight)
        Me.Controls.Add(Me.barDockControlBottom)
        Me.Controls.Add(Me.barDockControlTop)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "FRM__HHID"
        Me.Text = "IMPORT DATI DA FORN."
        CType(Me.dttSmartArt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NtsBarManager1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.fmTipoDato, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fmTipoDato.ResumeLayout(False)
        CType(Me.pnProgress, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnProgress.ResumeLayout(False)
        CType(Me.pbProgress.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pnPrincipale, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnPrincipale.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Public Overridable Sub InitControls()
        InitControlsBeginEndInit(Me, False)
        Try
            '-------------------------------------------------
            'carico le immagini della toolbar
            Try
                Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
                tlbZoom.Glyph = Bitmap.FromFile(oApp.ChildImageDir & "\zoom.png")
                tlbEsci.Glyph = Bitmap.FromFile(oApp.ChildImageDir & "\exit.png")
                tlbLog.Glyph = Bitmap.FromFile(oApp.ChildImageDir & "\Info_icon.png")
            Catch ex As Exception
                'non gestisco l'errore: se non c'è una immagine prendo quella standard
            End Try

            Try
                cmdImpAmazon.ImagePath = (oApp.ChildImageDir & "\Import-Orders_96.png")
                cmdImpAmazon.ImageLocation = DevExpress.XtraEditors.ImageLocation.MiddleCenter
            Catch ex As Exception
                'non gestisco l'errore: se non c'è una immagine prendo quella standard
            End Try
            '-------------------------------------------------
            'chiamo lo script per inizializzare i controlli caricati con source ext
            NTSScriptExec("InitControls", Me, Nothing)
        Catch ex As Exception
            '-------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '-------------------------------------------------
        End Try

        InitControlsBeginEndInit(Me, True)
    End Sub
    Public Overridable Sub Bindcontrols()
        Try
            '-------------------------------------------------
            'se i controlli erano già stati precedentemente collegati, li scollego
            NTSFormClearDataBinding(Me)
            '-------------------------------------------------
            'collego il BindingSource ai vari controlli 
            'edtb_dataHHWC.NTSDbField = "TABHHWC.tb_dataHHWC"
            'edtb_kmHHWC.NTSDbField = "TABHHWC.tb_kmHHWC"
            '-------------------------------------------------
            'per agganciare al dataset i vari controlli
            'NTSFormAddDataBinding(dcHHWC, Me)
        Catch ex As Exception
            '-------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '-------------------------------------------------
        End Try
    End Sub

#Region "Eventi Form"
    Private Sub FRM__HHID_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Try
            'dcHHWC.Dispose()
            'dsHHWC.Dispose()
        Catch
        End Try
    End Sub
    Private Sub FRM__HHID_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            'If Not Salva() Then e.Cancel = True
        Catch ex As Exception
            '-------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '-------------------------------------------------
        End Try
    End Sub
    Private Sub FRM__HHID_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
            '-------------------------------------------------
            'predispongo i controlli
            InitControls()

            'collego il BindingSource ai vari controlli 
            Bindcontrols()

            '-------------------------------------------------
            'sempre alla fine di questa funzione: applico le regole della gctl
            GctlSetRoules()

            '-------------------------------------------------
            'Leggo le opzioni di registro 
            oCleHhId.LeggiRegistroDoc()

        Catch ex As Exception
            '-------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '-------------------------------------------------
        Finally
            Me.Cursor = System.Windows.Forms.Cursors.Default
        End Try
    End Sub
#End Region

#Region "Eventi Button"
    Public Overridable Sub cmdImpAmazon_Click(sender As Object, e As EventArgs) Handles cmdImpAmazon.Click
        Dim strOut As String = String.Empty
        Try
            If ThreadCreaDoc IsNot Nothing Then ThreadCreaDoc.Abort()
        Catch ex As Exception
        End Try
        Try
            Control.CheckForIllegalCrossThreadCalls = False
            '----------------------------------------------------
            ThreadCreaDoc = New Threading.Thread(AddressOf ImportaDatiExcelAmazon)
            ThreadCreaDoc.SetApartmentState(Threading.ApartmentState.STA)
            ThreadCreaDoc.Start()
            ThreadCreaDoc.Join()
            ThreadCreaDoc.Abort()

            '*** Messaggi di fine esportazione ***
            If dsHhId IsNot Nothing AndAlso dsHhId.Tables.Contains("Corpo") Then
                strOut = "Importazione impegni da  terminata!" & vbCrLf &
                "- Righe processate: " & dsHhId.Tables("Corpo").Rows.Count & vbCrLf &
                "- Impegni generati: " & nCreati & vbCrLf &
                "- Errori generati: " & nErrori
                oApp.MsgBoxInfo(strOut)
            End If

        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
        End Try
    End Sub
    Public Overridable Sub cmdImpXbio_Click(sender As Object, e As EventArgs) Handles cmdImpXbio.Click
        Dim strOut As String = String.Empty
        Try
            If ThreadCreaDoc IsNot Nothing Then ThreadCreaDoc.Abort()
        Catch ex As Exception
        End Try
        Try
            Control.CheckForIllegalCrossThreadCalls = False
            '----------------------------------------------------
            ThreadCreaDoc = New Threading.Thread(AddressOf ImportaDatiExcelXbio)
            ThreadCreaDoc.SetApartmentState(Threading.ApartmentState.STA)
            ThreadCreaDoc.Start()
            ThreadCreaDoc.Join()
            ThreadCreaDoc.Abort()

            '*** Messaggi di fine esportazione ***
            If dsHhId IsNot Nothing AndAlso dsHhId.Tables.Contains("Corpo") Then
                strOut = "Importazione impegni da  terminata!" & vbCrLf &
                "- Righe processate: " & dsHhId.Tables("Corpo").Rows.Count & vbCrLf &
                "- Impegni generati: " & nCreati & vbCrLf &
                "- Errori generati: " & nErrori
                oApp.MsgBoxInfo(strOut)
            End If

        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
        End Try
    End Sub

    Private Sub cmdImpSleepH_Click(sender As Object, e As EventArgs) Handles cmdImpSleepH.Click
        Dim strOut As String = String.Empty
        Try
            If ThreadCreaDoc IsNot Nothing Then ThreadCreaDoc.Abort()
        Catch ex As Exception
        End Try
        Try
            Control.CheckForIllegalCrossThreadCalls = False
            '----------------------------------------------------
            ThreadCreaDoc = New Threading.Thread(AddressOf ImportaDatiExcelSleepH)
            ThreadCreaDoc.SetApartmentState(Threading.ApartmentState.STA)
            ThreadCreaDoc.Start()
            ThreadCreaDoc.Join()
            ThreadCreaDoc.Abort()

            '*** Messaggi di fine esportazione ***
            If dsHhId IsNot Nothing AndAlso dsHhId.Tables.Contains("Corpo") Then
                strOut = "Importazione impegni da  terminata!" & vbCrLf &
                "- Righe processate: " & dsHhId.Tables("Corpo").Rows.Count & vbCrLf &
                "- Impegni generati: " & nCreati & vbCrLf &
                "- Errori generati: " & nErrori
                oApp.MsgBoxInfo(strOut)
            End If

        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
        End Try
    End Sub
#End Region

#Region "Eventi Toolbar"
    Public Overridable Sub tlbZoom_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles tlbZoom.ItemClick
        Try
            NTSCallStandardZoom()
        Catch ex As Exception
            '-------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '-------------------------------------------------
        End Try
    End Sub
    Public Overridable Sub tlbEsci_ItemClick(ByVal sender As System.Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs) Handles tlbEsci.ItemClick
        Try
            Me.Close()
        Catch ex As Exception
            '-------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '-------------------------------------------------
        End Try
    End Sub
    Public Overridable Sub tlbLog_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles tlbLog.ItemClick
        Dim strFolderPath As String
        Dim strFileName As String

        Try
            strFolderPath = oApp.AscDir & "\BS__HHID\"
            strFileName = "Log_" & Now.Date.ToString("ddMMyyyy") & ".txt"

            If IO.File.Exists(strFolderPath & strFileName) Then
                Process.Start(strFolderPath & strFileName)
            ElseIf IO.Directory.Exists(strFolderPath) Then
                Process.Start(strFolderPath)
            Else
                oApp.MsgBoxInfo("Non sono presenti file di log!")
            End If
        Catch ex As Exception
            '-------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '-------------------------------------------------
        End Try
    End Sub
#End Region

    Public Overridable Function ImportaDatiExcelAmazon() As Boolean
        Dim strPath As String
        Dim strFoglio As String = "Foglio1"
        Dim strOut As String = String.Empty

        Dim oImpegni As New List(Of DTO_HHID_TestataImp)
        Try
            '--- Lettura File -----------------------------------------------------------------------------------------------------------------------
            '*** Faccio selezionare il file da importare ***
            strPath = TrovaFile()

            '*** Controlli validità doc. selezionato ***
            If Not oCleHhId.ControllaValiditaFile(strPath, strFoglio, strOut) Then
                oApp.MsgBoxInfo(strOut)
                Return False
            End If
            '--- Lettura Dati -----------------------------------------------------------------------------------------------------------------------
            '*** Leggo i dati nell'Testo ***
            If Not oCleHhId.getImpegniExcelAmazon(strPath, strFoglio, dsHhId) Then Return False

            '*** Controllo se c'è almeno un dato ***
            If dsHhId.Tables.Count = 0 OrElse Not dsHhId.Tables.Contains("Testa") Then GoTo Fine
            If dsHhId.Tables.Count = 0 OrElse Not dsHhId.Tables.Contains("Corpo") Then GoTo Fine

            oImpegni = oCleHhId.DataTableAmazonToDTO(dsHhId, oImpegni)
            ImportaDocumentiDaDTO(oImpegni)
Fine:
            '*** Reset progress bar ***
            pbProgress.EditValue = 0
            lbProgress.Text = "0/0"

            ScriviLog("")
            Return True
        Catch ex As Exception
            '---------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '---------------------------------------------------------
        End Try
    End Function

    Public Overridable Function ImportaDatiExcelXbio() As Boolean
        Dim strPath As String
        Dim strFoglio As String = "Foglio1"
        Dim strOut As String = String.Empty

        Dim oImpegni As New List(Of DTO_HHID_TestataImp)
        Try
            '--- Lettura File -----------------------------------------------------------------------------------------------------------------------
            '*** Faccio selezionare il file da importare ***
            strPath = TrovaFile()

            '*** Controlli validità doc. selezionato ***
            If Not oCleHhId.ControllaValiditaFile(strPath, strFoglio, strOut) Then
                oApp.MsgBoxInfo(strOut)
                Return False
            End If
            '--- Lettura Dati -----------------------------------------------------------------------------------------------------------------------
            '*** Leggo i dati nell'Testo ***
            If Not oCleHhId.getImpegniExcelXbio(strPath, strFoglio, dsHhId) Then Return False

            '*** Controllo se c'è almeno un dato ***
            If dsHhId.Tables.Count = 0 OrElse Not dsHhId.Tables.Contains("Testa") Then GoTo Fine
            If dsHhId.Tables.Count = 0 OrElse Not dsHhId.Tables.Contains("Corpo") Then GoTo Fine

            oImpegni = oCleHhId.DataTableXbioToDTO(dsHhId, oImpegni)
            ImportaDocumentiDaDTO(oImpegni)
Fine:
            '*** Reset progress bar ***
            pbProgress.EditValue = 0
            lbProgress.Text = "0/0"

            ScriviLog("")
            Return True
        Catch ex As Exception
            '---------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '---------------------------------------------------------
        End Try
    End Function

    Public Overridable Function ImportaDatiExcelSleepH() As Boolean
        Dim strPath As String
        Dim strFoglio As String = "Foglio1"
        Dim strOut As String = String.Empty

        Dim oImpegni As New List(Of DTO_HHID_TestataImp)
        Try
            '--- Lettura File -----------------------------------------------------------------------------------------------------------------------
            '*** Faccio selezionare il file da importare ***
            strPath = TrovaFile()

            '*** Controlli validità doc. selezionato ***
            If Not oCleHhId.ControllaValiditaFile(strPath, strFoglio, strOut) Then
                oApp.MsgBoxInfo(strOut)
                Return False
            End If
            '--- Lettura Dati -----------------------------------------------------------------------------------------------------------------------
            '*** Leggo i dati nell'Testo ***
            If Not oCleHhId.getImpegniExcelSleepH(strPath, strFoglio, dsHhId) Then Return False

            '*** Controllo se c'è almeno un dato ***
            If dsHhId.Tables.Count = 0 OrElse Not dsHhId.Tables.Contains("Testa") Then GoTo Fine
            If dsHhId.Tables.Count = 0 OrElse Not dsHhId.Tables.Contains("Corpo") Then GoTo Fine

            oImpegni = oCleHhId.DataTableSleepHToDTO(dsHhId, oImpegni)
            ImportaDocumentiDaDTO(oImpegni)
Fine:
            '*** Reset progress bar ***
            pbProgress.EditValue = 0
            lbProgress.Text = "0/0"

            ScriviLog("")
            Return True
        Catch ex As Exception
            '---------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '---------------------------------------------------------
        End Try
    End Function
#Region "FunzioniImport"
    Public Overridable Function ImportaDocumentiDaDTO(ByVal oImpegni As List(Of DTO_HHID_TestataImp)) As Boolean
        Dim strOut As String = String.Empty
        Dim bArticoloDef As Boolean

        Dim dsGsor As New DataSet
        Try
            nCreati = 0
            nDuplicati = 0
            nErrori = 0
            '--- Leggo dati dal database --------------------------------------------------------------------------------------------------------------
            dsGsor.Tables.Clear()
            dsGsor.Tables.Add(oCleHhId.oCldHhId.OpenRecordset("Select * FROM testord where td_riferim <> '' ", CLE__APP.DBTIPO.DBAZI, "TESTORD").Tables("TESTORD").Copy)
            '--- Imposta ProgressBar ------------------------------------------------------------------------------------------------------------------
            pbProgress.EditValue = 0
            pbProgress.Properties.Minimum = 0
            pbProgress.Properties.Maximum = oImpegni.Count
            pbProgress.Properties.Step = 1
            lbProgress.Text = "0/" & pbProgress.Properties.Minimum
            '--- Creazione Entity Esterne -------------------------------------------------------------------------------------------------------------
            InizializzaBEORGSOR()
            InizializzaBE__CLIE()
            '--- Ciclo di creazione degli DOCUMENTI ---------------------------------------------------------------------------------------------------
            For Each oImpegno As DTO_HHID_TestataImp In oImpegni
                '*** Comunico che sto per importare l'impegno ***
                ScriviLog("--- Importazione Impegno: " & NTSCStr(oImpegno.strRiferim).Trim & " ---")

                '*** Controllo se il documento già esiste ***
                If dsGsor.Tables("TESTORD").Select("td_riferim = " & CStrSQL(NTSCStr(oImpegno.strRiferim.Trim))).Length = 0 Then
                    '*** Imposto i dati base dell'impegno ***
                    oImpegno.strTipoOrd = "R"
                    oImpegno.nAnnoOrd = oImpegno.dDatDoc.Year
                    'oImpegno.strSerieOrd = " "
                    oImpegno.lNumTmpOrd = 0

                    '----------------------------------------------------------------------------------------
                    '--- Legge il progressivo in TABNUMA
                    '----------------------------------------------------------------------------------------
                    'oImpegno.strSerieOrd = oCleHhId.strSerie_Impegni
                    oImpegno.lNumTmpOrd = oCleGsor.LegNuma(oImpegno.strTipoOrd, oImpegno.strSerieOrd, oImpegno.nAnnoOrd)
                    If oImpegno.lNumTmpOrd = 0 Then
                        ScriviLog($"Controllare le numerazioni! Da creare ({oImpegno.strTipoOrd} - Impegni, serie: {oImpegno.strSerieOrd}, anno: {oImpegno.nAnnoOrd})")
                        Continue For
                    End If
                    '----------------------------
                    'preparo l'ambiente
                    Dim ds As New DataSet
                    If Not oCleGsor.ApriOrdine(oApp.Ditta, False, oImpegno.strTipoOrd, oImpegno.nAnnoOrd, oImpegno.strSerieOrd, oImpegno.lNumTmpOrd, ds) Then Return False
                    oCleGsor.bInApriDocSilent = True
                    If oCleGsor.dsShared.Tables("TESTA").Rows.Count > 0 Then
                        'ordine già esistente!!!!!!!!
                        ScriviLog("Controllare le numerazioni! Sono disallineate")
                        Return False
                    End If

                    '*** Abilito/Disabilito alcuni controlli ***
                    oCleGsor.bDisabilitaCheckDateAnteriori = True

                    Try
                        '*** Creo il nuovo documento ***
                        oCleGsor.ResetVar()
                        oCleGsor.strVisNoteConto = "N"
                        oCleGsor.NuovoOrdine(oApp.Ditta, oImpegno.strTipoOrd, oImpegno.nAnnoOrd, oImpegno.strSerieOrd, oImpegno.lNumTmpOrd, "")
                        oCleGsor.bInNuovoDocSilent = True

                        '*** Creo la testata del documento ***
                        If Not oCleHhId.CreaTestataOrd(oImpegno, oCleClie, oCleGsor) Then
                            Throw New NTSException(oApp.Tr(Me, 128607611686875000, "Errore sulla testata dell'impegno!"))
                        End If

                        '*** Creo il corpo del documento ***
                        bArticoloDef = False
                        If Not oCleHhId.CreaRigheOrd(oImpegno, oCleGsor, bArticoloDef) Then
                            Throw New NTSException(oApp.Tr(Me, 128607611686875000, "Errore sul corpo dell'impegno!"))
                        End If
                        '*** Se ho usato l'articolo descrittivo perchè non l'ho trovato lo segnalo ***
                        If bArticoloDef Then nErrori += 1

                        '*** Creo il piede del documento ***
                        If Not oCleHhId.SettaPiedeOrd(oImpegno, oCleGsor) Then
                            Throw New NTSException(oApp.Tr(Me, 128607611686875000, "Errore sul piede dell'impegno!"))
                        End If

                        '*** Salvo le modifiche ***
                        If Not oCleGsor.SalvaOrdine("N") Then 'U = update, D = delete
                            Throw New NTSException(oApp.Tr(Me, 128607611686875000, "Errore al salvataggio dell'impegno! riferimento ordine: " & vbCrLf & "|" & NTSCStr(oImpegno.strRiferim) & "|"))
                        Else
                            oCleGsor.NuovoDocDaImportExport(oApp.Ditta, oImpegno.strTipoOrd, oImpegno.nAnnoOrd, oImpegno.strSerieOrd, 0, True)
                            nCreati += 1
                        End If

                    Catch ex As Exception
                        nErrori += 1
                        ScriviLog($"Errore generazione impegno: {NTSCStr(oImpegno.strRiferim)}")
                        ScriviLog(ex.Message)
                    End Try
                Else
                    ScriviLog("Il documento esiste già su Business quindi non verrà importato")
                    nDuplicati += 1
                End If

                '*** Avanzamento della ProgressBar ***
                pbProgress.PerformStep()
                lbProgress.Text = "Impegni: " & pbProgress.Text & "/" & pbProgress.Properties.Maximum
            Next
Fine:
            '*** Reset progress bar ***
            pbProgress.EditValue = 0
            lbProgress.Text = "0/0"

            ScriviLog("")
            Return True
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
        End Try
    End Function
#End Region

#Region "Utility"
    Public Overridable Function TrovaFile() As String
        Dim strResult As String
        Try
            '-----------------------------------------------------------------------------------------
            '--- Adesso apre la Common Dialog sulla cartella delle immagini
            '-----------------------------------------------------------------------------------------
            If OFD_File Is Nothing Then OFD_File = New NTSOpenFileDialog
            OFD_File.CheckFileExists = True
            OFD_File.ShowReadOnly = False
            OFD_File.ShowHelp = False
            OFD_File.DefaultExt = "xlsx"
            OFD_File.Title = "Selezione file excel"
            OFD_File.Filter = "Excel Files|*.xlsx;*.xls;*.xlsm|CSV Files|*.csv"
            OFD_File.InitialDirectory = oApp.ServerDir
            OFD_File.FileName = ""
            OFD_File.oMenu = oMenu
            OFD_File.ShowDialog()

            '-----------------------------------------------------------------------------------------
            If OFD_File.FileName.EndsWith(".csv") Then
                strResult = oCleHhId.ConvertCSVtoXLSX(OFD_File.FileName)
            Else
                strResult = OFD_File.FileName
            End If
            '-----------------------------------------------------------------------------------------
            Return strResult
            '---------------------------------------------------------------------------------------
        Catch ex As Exception
            '-------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            Return ""
            '-------------------------------------------------
        End Try
    End Function
    Public Overridable Sub ScriviLog(ByVal strText As String)
        Try
            oCleHhId.ScriviLog(strText)
        Catch ex As Exception
            CLN__STD.GestErr(ex, Me, "")
        End Try
    End Sub
#End Region

#Region "Creazione Entity Esterne"
    Public oCleGsor As CLEORGSOR = Nothing
    Public Overloads Function InizializzaBEORGSOR() As Boolean
        Try
            '------------------------------------------------
            'creo e attivo l'entity e inizializzo la funzione che dovrà rilevare gli eventi dall'ENTITY
            Dim strErr As String = ""
            Dim oTmp As Object = Nothing
            If CLN__STD.NTSIstanziaDll(oApp.ServerDir, oApp.NetDir, "BN__HHID", "BEORGSOR", oTmp, strErr, False, "", "") = False Then
                oApp.MsgBoxErr(oApp.Tr(Me, 128550728307822408, "ERRORE in fase di creazione Entity:") & vbCrLf & strErr)
                Return False
            End If
            oCleGsor = CType(oTmp, CLEORGSOR)
            '------------------------------------------------

            'AddHandler oCleGsor.RemoteEvent, AddressOf GestisciEventiEntity
            If oCleGsor.Init(oApp, oScript, oMenu.oCleComm, "", False, "", "") = False Then Return False
            oCleGsor.InitExt()

            Return True
        Catch ex As Exception
            '---------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '---------------------------------------------------------
        End Try

    End Function
    Public oCleClie As CLE__CLIE = Nothing
    Public Overridable Function InizializzaBE__CLIE() As Boolean
        Try
            If Not oCleClie Is Nothing Then Return True
            '------------------------
            'inizializza entity
            Dim strErr As String = ""
            Dim oTmp As Object = Nothing
            If CLN__STD.NTSIstanziaDll(oApp.ServerDir, oApp.NetDir, "BN__CLIE", "BE__CLIE", oTmp, strErr, False, "", "") = False Then
                oApp.MsgBoxErr(oApp.Tr(Me, 128271029889882656, "ERRORE in fase di creazione Entity:" & vbCrLf & "|" & strErr & "|"))
                Return False
            End If
            oCleClie = CType(oTmp, CLE__CLIE)
            '------------------------------------------------

            'AddHandler oCleClie.RemoteEvent, AddressOf GestisciEventiEntity
            If oCleClie.Init(oApp, Nothing, oMenu.oCleComm, "", False, "", "") = False Then Return False

            Return True
        Catch ex As Exception
            '---------------------------------------------------------
            CLN__STD.GestErr(ex, Me, "")
            '---------------------------------------------------------
        End Try
    End Function
#End Region
End Class