Imports System.Runtime.Remoting
Imports System.Runtime.Remoting.Channels
Imports System.Runtime.Remoting.Channels.Tcp
Imports System.Security.Cryptography.X509Certificates
Imports System.IO
Imports System.Configuration
Imports System.Reflection
Imports System.Threading.Thread
Imports System.Globalization
Imports Microsoft.VisualBasic
Imports Microsoft.Win32
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Security.Cryptography
Imports System.Data

Public Class FRMBUSCUBE

  '************** RICORDARSI DI COMPILARE SEMPRE CON L'OPZIONE CPU x86, ALTRIMENTI NON GIRA LA CREAZIONE DEGLI ACTIVEX DI BUSINESS SU 64 BIT!!! ***********  compila->opzioni di compilazione avanzate->cpu di destinazione=x86

  'quando si cambia l'icona di business deve essere fatto anche in bn__Chil e bn__std come icona collegata alla form!!!!

  'per mutex
  'Costanti nome dei MUTEX: a seguire il nome della dir dove cira l'EXE, così il mutex diventa specifico per directory
  Dim bsMTXBus As String = "NtsBusinessMutex" 'BUSINESS, BUSBATCH
  Dim bsMTXUtl As String = "NtsBusUtilMutex" 'BUSADMINTOOLS
  Dim bsMTXSet As String = "NtsBusSetupMutex" 'BUSSETUP
  Dim bsMTXAgg As String = "NtsBusAggMutex" 'BUSAGG
  Dim bsMTXSrv As String = "NtsBusServMutex" 'BUSSERV
  Dim oMutexG As System.Threading.Mutex = Nothing
  Dim oMutexL As System.Threading.Mutex = Nothing

  Public bBatch As Boolean = False
  Public DirP As String = System.Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
  Public LogDirP As String = ""
  Public strParam As String = ""

  Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Integer) As Integer
  Private Declare Function SetSecurityDescriptorDacl Lib "advapi32.dll" (ByVal pSecurityDescriptor As IntPtr, ByVal bDaclPresent As Integer, ByVal pDacl As Byte, ByVal bDaclDefaulted As Integer) As Integer
  Private Declare Function InitializeSecurityDescriptor Lib "advapi32.dll" (ByVal pSecurityDescriptor As IntPtr, ByVal dwRevision As Integer) As Integer
  Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, ByVal lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByVal Args As Long) As Long

  Private Const MAX_PATH As Long = 260
  Private Const LB_SETTABSTOPS As Long = &H192
  Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000
  Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200
  Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK As Long = &HFF
  Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY As Long = &H2000

  Structure SECURITY_ATTRIBUTES
    Dim nLength As Integer
    Dim lpSecurityDescriptor As Integer
    Dim bInheritHandle As Integer
  End Structure
  Private Structure ACL
    Dim AclRevision As Byte
    Dim Sbz1 As Byte
    Dim AclSize As Integer
    Dim AceCount As Integer
    Dim Sbz2 As Integer
  End Structure
  Private Structure SECURITY_DESCRIPTOR
    Dim Revision As Byte
    Dim Sbz1 As Byte
    Dim Control As Integer
    Dim Owner As Integer
    Dim Group As Integer
    Dim sACL As ACL
    Dim Dacl As ACL
  End Structure
  Private Const SECURITY_DESCRIPTOR_REVISION = 1
  Private Const ERROR_ALREADY_EXISTS = 183&
  Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
  Private Const SYNCHRONIZE = &H100000
  Private Const MUTANT_QUERY_STATE = &H1
  Private Const MUTEX_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or MUTANT_QUERY_STATE

  Const LOGON_WITH_PROFILE As Integer = &H1
  Const CREATE_DEFAULT_ERROR_MODE As Integer = &H4000000

  Public Structure PROCESS_INFO
    Public hProcess As IntPtr
    Public hThread As IntPtr
    Public dwProcessId As Integer
    Public dwThreadId As Integer
  End Structure
  Public Structure STARTUP_INFO
    Public cb As Integer
    Public lpReserved As Integer
    <MarshalAs(UnmanagedType.LPTStr)> Public lpDesktop As String
    <MarshalAs(UnmanagedType.LPTStr)> Public lpTitle As String
    Public dwX As Long
    Public dwY As Integer
    Public dwXSize As Integer
    Public dwYSize As Integer
    Public dwXCountChars As Integer
    Public dwYCountChars As Integer
    Public dwFillAttribute As Integer
    Public dwFlags As Integer
    Public wShowWindow As Short
    Public cbReserved2 As Short
    Public lpReserved2 As Integer
    Public hStdInput As Integer
    Public hStdOutput As Integer
    Public hStdError As Integer
  End Structure
  <DllImport("advapi32.dll")>
  Public Shared Function CreateProcessWithLogonW(<MarshalAs(UnmanagedType.LPWStr)> ByVal lpUsername As String,
                                          <MarshalAs(UnmanagedType.LPWStr)> ByVal lpDomain As String,
                                          <MarshalAs(UnmanagedType.LPWStr)> ByVal lpPassword As String,
                                          ByVal dwLogonFlags As Integer,
                                          <MarshalAs(UnmanagedType.LPWStr)> ByVal lpApplicationName As String,
                                          <MarshalAs(UnmanagedType.LPWStr)> ByVal lpCommandLine As String,
                                          ByVal lpCreationFlags As Integer,
                                          ByVal lpVoid As Integer,
                                          <MarshalAs(UnmanagedType.LPWStr)> ByVal lpCurrentDirectory As String,
                                          ByRef lpStartupInfo As STARTUP_INFO,
                                          ByRef lpProcessInfo As PROCESS_INFO) As Integer
  End Function
  Private Declare Function GetLastError Lib "kernel32" () As Long

  Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Integer, ByRef phiconLarge As Integer, ByRef phiconSmall As Integer, ByVal nIcons As Integer) As Integer

  Private oStd As Object = Nothing

  Public Enum Process_DPI_Awareness
    Process_DPI_Unaware = 0
    Process_System_DPI_Aware = 1
    Process_Per_Monitor_DPI_Aware = 2
  End Enum
  <DllImport("shcore.dll")>
  Public Shared Function SetProcessDpiAwareness(value As Process_DPI_Awareness) As Integer
  End Function
  <DllImport("SHCore.dll", SetLastError:=True)>
  Public Shared Sub GetProcessDpiAwareness(hprocess As IntPtr, ByRef awareness As Process_DPI_Awareness)
  End Sub

  Sub New()

    '*************************************************************
    'carico BN__STD
    Try
      Dim assem1 As System.Reflection.Assembly
      Dim tpChild1 As Type

      assem1 = System.Reflection.Assembly.LoadFrom(Application.StartupPath & "\BN__STD.DLL")
      tpChild1 = assem1.GetType("NTSInformatica.CLN__STD", True, True)
      oStd = Activator.CreateInstance(tpChild1)

    Catch ex As Exception
      If bBatch = False Then
        MsgBox("Error starting Bus.net: " & ex.Message, MsgBoxStyle.Critical, "Messaggio di Business CUBE")
      Else
        WriteMsgBoxToLog("Error starting Bus.net: " & ex.ToString, True)
      End If
      Me.Close()
      Return
    End Try

    'se sono in terminal server imposto DPI AWARE = False, altrimenti non funziona niente!!!!!
    'diversamente imposto DPI AWARE = TRUE/PM.
    'occhio: non va settato nel manifest, perchè quello avrebbe la precedenza!!!!
    Try
      Dim n1 As Process_DPI_Awareness
      GetProcessDpiAwareness(Process.GetCurrentProcess.Handle, n1)
      If oStd.IsTerminalServer Then
        Dim nRes As Integer = SetProcessDpiAwareness(Process_DPI_Awareness.Process_DPI_Unaware)
        'MsgBox("NO")
      Else
        Dim nRes As Integer = SetProcessDpiAwareness(Process_DPI_Awareness.Process_Per_Monitor_DPI_Aware)
        'MsgBox("PM")
      End If
      Dim nErr As Integer = Marshal.GetLastWin32Error()
    Catch ex As Exception
      'per default è no DPI-AWARE!!! come net 2015
    End Try

    InitializeComponent()
  End Sub

  Private Sub fmBusnet_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    'modifiche per windows 10 modalità TABLET (compatibili anche non in modalità tablet!!!)
    'per ora l'unico limite è che se, in modalità tablet, faccio partire bus da taskbar qundo è in visualizzazione il menu di windows
    'busnet parte ma rimane nascosto dal menu. lo stesso problema c'è anche coon notepad++ mentre winword no...

    'oltre alle righe qui sotto sono state aggiunte le .BringToFront
    'e in BN__MENU l'unica form cha ha 'showintaskbar' è quella della scelta operatore

    Me.Height = 0
    Me.Width = 0
    Me.MaximumSize = New Size(1, 1)
    Me.Size = Me.MaximumSize
    Me.Opacity = 0
    Me.FormBorderStyle = FormBorderStyle.SizableToolWindow  'così non si vede nell'alt+tab di windows

    'prendo l'icona dall'exe
    Dim bigIcon As Icon = Nothing
    Dim smallIcon As Icon = Nothing
    Me.Icon = ReturnIcon(Application.ExecutablePath, 0)
    'Me.IsMdiContainer = True per abilitare ambiente MDI

    bsMTXBus += "_" & Application.StartupPath.Replace(":", "_").Replace("\", "_").Replace("(", "_").Replace(")", "_").Replace(" ", "_").ToUpper
    bsMTXUtl += "_" & Application.StartupPath.Replace(":", "_").Replace("\", "_").Replace("(", "_").Replace(")", "_").Replace(" ", "_").ToUpper
    bsMTXSet += "_" & Application.StartupPath.Replace(":", "_").Replace("\", "_").Replace("(", "_").Replace(")", "_").Replace(" ", "_").ToUpper
    bsMTXAgg += "_" & Application.StartupPath.Replace(":", "_").Replace("\", "_").Replace("(", "_").Replace(")", "_").Replace(" ", "_").ToUpper
    bsMTXSrv += "_" & Application.StartupPath.Replace(":", "_").Replace("\", "_").Replace("(", "_").Replace(")", "_").Replace(" ", "_").ToUpper
  End Sub

  Private Function ReturnIcon(ByVal Path As String, ByVal Index As Integer, Optional ByVal small As Boolean = False) As Icon
    'estrae l'icona dall'eseguibile
    Dim bigIcon As Integer
    Dim smallIcon As Integer
    ExtractIcon(Path, Index, bigIcon, smallIcon, 1)
    If bigIcon <> 0 Then
      If small = False Then
        Return Icon.FromHandle(bigIcon)
      Else
        Return Icon.FromHandle(smallIcon)
      End If
    Else
      Return Nothing
    End If
  End Function

  Private Sub fmBusnet_Shown(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Shown
    'Dim strTmp As String = ""
    Dim strDir As String
    Dim i As Integer
    Dim bDisableMutex As Boolean = False
    Dim bExists As Boolean = False
    Dim strProfilo As String = ""

    'Verifico che non sia presente l'impostazione per usare UTF-8 beta
    'Se presente, emette messaggio di avviso ed esco
    Dim oAcp As Object = My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Nls\CodePage", "ACP", Nothing)
    Dim oMaccp As Object = My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Nls\CodePage", "MAACP", Nothing)
    Dim oOemcp As Object = My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Nls\CodePage", "OEMCP", Nothing)

    If (Not oAcp Is Nothing AndAlso oAcp.ToString() = "65001") Or (Not oMaccp Is Nothing AndAlso oMaccp.ToString() = "65001") Or (Not oOemcp Is Nothing AndAlso oOemcp.ToString() = "65001") Then
      If bBatch = False Then
        MsgBox("Rilevata l'impostazione per usare UTF-8 beta, impossibile avviare Business.")
      Else
        WriteMsgBoxToLog("Rilevata l'impostazione per usare UTF-8 beta, impossibile avviare Business.", True)
      End If
      Me.Close()
      Return
    End If

    'strDir = Directory.GetCurrentDirectory
    strDir = Application.StartupPath

    If My.Application.CommandLineArgs.Count > 0 Then
      For i = 0 To My.Application.CommandLineArgs.Count - 1
        strParam = strParam & My.Application.CommandLineArgs(i).ToString.Replace(" ", "*§§§*§§§*") & " "
      Next
    End If
    If strParam = "" Then
      'a volte, in debug, non legge la riga di comando ...
      strParam = ". . . Business"
    Else
      strParam = Trim(strParam)
    End If

    'imposto la dir in bn__std
    strProfilo = CheckBusProfilo(Command, False)
    DirP = oStd.GetSettingReg("Business", strProfilo, "Dir", "")
    If DirP = "*-*" Then
      'non è stato trovato il file .cfg o la stringa di connessione è sbagliata: chiudo tutto
      Me.Close()
      Return
    End If
    If Not System.IO.Directory.Exists(DirP) Then
      DirP = System.Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
    End If
    'Crea la cartella dei log: Asc\Log\ANNO\MESE
    LogDirP = oStd.GetSettingReg("Business", strProfilo, "Dir", "") & "\Asc\Log\" & Now.Year & "\" & Now.Month.ToString("00")
    If Not Directory.Exists(LogDirP) Then Directory.CreateDirectory(LogDirP)

    If strParam.ToUpper.IndexOf(" /B ") > -1 Then
      bBatch = True
      Dim strParamCrypt As String = ""
      Dim strParams As String() = strParam.Split(" "c)
      If strParams.Length >= 2 Then
        strParams(1) = "***"
        For Each strParametro As String In strParams
          strParamCrypt += strParametro + " "
        Next
        WriteMsgBoxToLog("Busness NET Avviato con la seguente riga di comando: " & strParamCrypt, False)
      Else
        WriteMsgBoxToLog("Busness NET Avviato con la seguente riga di comando: " & strParam, False)
      End If
    End If

    '*************************************************************
    'PER PRIMA COSA IMPOSTO LA DATA NEL FORMAT OCON ANNO A 4 CIFRE, DIVERSAMENTE SE INSERISCO UNA DATA NEL 
    'FORMATO 31/12/2099 VIENE TRADOTTA IN 31/12/1999 QUANDO SI FA .toshortdatestring!!!!!
    Try
      Dim strDateFormat As String = CurrentThread.CurrentCulture.DateTimeFormat.ShortDatePattern
      If strDateFormat.IndexOf("yyyy") = -1 Then
        strDateFormat = strDateFormat.Replace("yy", "yyyy")
        Dim oCulture As CultureInfo = New CultureInfo(CurrentThread.CurrentCulture.Name)
        oCulture.DateTimeFormat.ShortDatePattern = strDateFormat
        CurrentThread.CurrentCulture = oCulture
        'strTmp = CurrentThread.CurrentCulture.DateTimeFormat.ShortDatePattern
      End If
    Catch ex As Exception
      If bBatch = False Then
        MsgBox("Error starting Bus.net during set DATETIME format: " & ex.ToString)
      Else
        WriteMsgBoxToLog("Error starting Bus.net during set DATETIME format: " & ex.ToString, True)
      End If
      Me.Close()
      Return
    End Try

    '*************************************************************
    'preparo il canale per la comunicazione dei REMOTING:
    'dentro ai vari entity ci sarà una riga del tipo quella commentata per avvisare che il dal non è in locale ma è in remoto

    'Controlla se sono disabilitati i Mutex
    bDisableMutex = CBool(Val(oStd.GetSettingReg("Business", "Debug", "DisableMutex", "0")))
    If Not bDisableMutex Then
      'Controlla che nonsiano in esecuzione programmi di utilità (BusAdminTools)
      If NtsMutexTrova(bsMTXUtl) Then
        If bBatch = False Then
          MsgBox("Il programma BusAdminTools.exe è in esecuzione. Chiuderlo prima di eseguire l'applicazione.", vbExclamation, "Messaggio di Business CUBE")
        Else
          WriteMsgBoxToLog("Il programma BusAdminTools.exe è in esecuzione. Chiuderlo prima di eseguire l'applicazione.", True)
        End If
        Me.Close()
        Return
      End If
      'Controlla che non ci siano in esecuzione BUSAGG
      If NtsMutexTrova(bsMTXAgg) Then
        If bBatch = False Then
          MsgBox("Uno o più componenti di Business, sotto indicati, sono attualmente in esecuzione:" & vbCrLf & vbCrLf & "BUSAGG.EXE (Installazione aggiornamenti)" & vbCrLf & vbCrLf & "Chiudere tutte le istanze dei programmi prima di eseguire l'applicazione.", vbExclamation, "Messaggio di Business CUBE")
        Else
          WriteMsgBoxToLog("Uno o più componenti di Business, sotto indicati, sono attualmente in esecuzione:" & vbCrLf & vbCrLf & "BUSAGG.EXE (Installazione aggiornamenti)" & vbCrLf & vbCrLf & "Chiudere tutte le istanze dei programmi prima di eseguire l'applicazione.", True)
        End If
        Me.Close()
        Return
      End If
      'Controlla che non ci siano in esecuzione BUSSETUP
      If NtsMutexTrova(bsMTXSet) Then
        If bBatch = False Then
          MsgBox("Uno o più componenti di Business, sotto indicati, sono attualmente in esecuzione:" & vbCrLf & vbCrLf & "BUSSETUP.EXE (Installazione di Business) " & vbCrLf & vbCrLf & "Chiudere tutte le istanze dei programmi prima di eseguire l'applicazione.", vbExclamation, "Messaggio di Business CUBE")
        Else
          WriteMsgBoxToLog("Uno o più componenti di Business, sotto indicati, sono attualmente in esecuzione:" & vbCrLf & vbCrLf & "BUSSETUP.EXE (Installazione di Business) " & vbCrLf & vbCrLf & "Chiudere tutte le istanze dei programmi prima di eseguire l'applicazione.", True)
        End If
        Me.Close()
        Return
      End If
    End If

    If oStd.GetNoLogin(UCase(strProfilo)) Then
      If bBatch = False Then
        Dim strFile As String = Path.GetTempPath() & "Th2_" & System.Diagnostics.Process.GetCurrentProcess.Id & ".txt"
        Dim w1 As New System.IO.StreamWriter(strFile, False)
        w1.Write("Messaggio da Business Cube:" & vbCrLf & vbCrLf & "ATTENZIONE: è stato impostato per non accettare nuove connessioni. Impossibile proseguire.")
        w1.Close()
        Dim p As New System.Diagnostics.Process()
        p.StartInfo.FileName = Application.StartupPath & "\buschat.exe"
        p.StartInfo.Arguments = """" & strFile & """" 'nome del file
        p.StartInfo.CreateNoWindow = False
        p.Start()
      End If
      Me.Close()
      Return
    End If

    'esco se l'aggiornamento non deve partire in automatico all'avvio di Business
    If bBatch = False Then
      If oStd.GetSettingReg("BUSINESS", UCase(strProfilo) & "\BUSAGG", "BusaggAutoUpdate", "N") <> "N" Then
        'lancio la routine che verifica se ci sono aggiornamenti da caricare prima di avviare Busiess
        If Not bDisableMutex Then
          'L'installazione aggiornamenti la fa partire solo se non è già in esecuzione Business
          If Not NtsMutexTrova(bsMTXBus) Then
            If Not CaricaAgg() Then
              Me.Close()
              Return
            End If
          End If
        Else
          If Not CaricaAgg() Then
            Me.Close()
            Return
          End If
        End If
      End If    'If strTmp <> "N" Then
    End If    'If bBatch = False Then

    If Not bDisableMutex Then
      'Crea il mutex per il controllo se l'applicazione è in esecuzione
      'DOPO AVER INSTALLATO GLI AGGIORNAMENTI
      NtsMutexCrea(bsMTXBus, bExists)
    End If

    '*************************************************************
    'carico l'oggetto MENU e lo avvio passandogli la riga di comando
    Try
      Dim assem As System.Reflection.Assembly
      Dim tpChild As Type
      Dim fmMenu As Object

      assem = System.Reflection.Assembly.LoadFrom(strDir & "\BN__MENU.DLL")
      tpChild = assem.GetType("NTSInformatica.FRMMENU", True, True)
      fmMenu = Activator.CreateInstance(tpChild)

      If Not fmMenu.InitMenu(strParam, strDir, Me, Nothing, DirP) Then
        Me.Close()
        Return
      End If

      Me.BringToFront()

      'visualizzo il menu (non modale)
      fmMenu.show()
      fmMenu.BringToFront

    Catch ex As Exception
      If bBatch = False Then
        MsgBox("Error starting Bus.net: " & ex.Message, MsgBoxStyle.Critical, "Messaggio di Business CUBE")
      Else
        WriteMsgBoxToLog("Error starting Bus.net: " & ex.ToString, True)
      End If
      Me.Close()
      Return
    End Try
  End Sub

  Private Sub fmBusnet_Unload(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
    'Elimina il mutex che eventualmente ha creato all'avvio
    Try
      NtsMutexElimina()
    Catch ex As Exception

    End Try
  End Sub


  Public Overridable Function CaricaAgg() As Boolean
    'carico gli aggiornamenti all'avvio di busnet
    'Dim objAgg As Object
    Dim strBusAggdir As String
    Dim strProfilo As String
    Dim bUserXXX As Boolean = False
    Dim strAggUser As String = ""
    Dim strAggPwd As String = ""
    Dim strDomain As String = ""
    Dim strUser As String = ""
    Dim nAggNumberClient As Integer = 0
    Dim nAggNumberServer As Integer = 0
    Dim i As Integer = 0
    'Dim s As New Security.SecureString

    Try
      strProfilo = CheckBusProfilo(Command, bUserXXX)

      'leggo le dir dove dovrebbero esserci gli aggiornamenti

      strBusAggdir = oStd.GetSettingReg("BUSINESS", UCase(strProfilo) & "\BUSAGG", "BusAggDir", "")
      If strBusAggdir.Length > 0 Then
        If strBusAggdir.Substring(strBusAggdir.Length - 1) <> "\" Then strBusAggdir += "\"
      End If

      'versione di Business locale
      nAggNumberServer = Val(oStd.GetSettingReg("BUSINESS", UCase(strProfilo) & "\BUSAGG", "AggNumber", "0"))

      'versione di Business sul server
      'Controlla se il Server è raggiungibile
      'e se esiste il file .txt dove è salvato il numero di aggiornamento nel Server
      nAggNumberClient = nAggNumberServer
      If System.IO.File.Exists(strBusAggdir & "Aggnumber.txt") = False Then
        MsgBox("Non esiste il file '" & strBusAggdir & "Aggnumber.txt'," & vbCrLf & "o il Server non è raggiungibile.", vbExclamation)
        Return True
      Else
        Dim r1 As New System.IO.StreamReader(strBusAggdir & "Aggnumber.txt")
        If Not r1.EndOfStream Then
          nAggNumberClient = Val(r1.ReadLine)
        End If
        r1.Close()
      End If

      'esco se le versioni sono uguali o se il numero presente nel file Aggnumber.txt è minore(WSS N° 247435)
      If nAggNumberClient <= nAggNumberServer Then Return True

      '-----------------------------------------
      'devo eseguire l'aggiornamento: se posso carico la versione aggiornata di busagg.exe
      'da NET 2013 il busagg.exe + busagm.exe sono sempre nella cartella \\server\bus\agg!!!!
      strAggUser = oStd.GetSettingReg("BUSINESS", UCase(strProfilo) & "\BUSAGG", "BusAggUser", "")
      strAggPwd = oStd.GetSettingReg("BUSINESS", UCase(strProfilo) & "\BUSAGG", "BusAggPwd", "")
      If Not System.IO.File.Exists(strBusAggdir & "Busagg.exe") Then
        MsgBox("Attenzione: non è possibile avviare la procedura di aggiornamento in quanto non è stato trovato il file '" & strBusAggdir & "Busagg.exe'", vbExclamation)
        Return True
      End If

      If Not VerifyFileSignature(strBusAggdir & "Busagg.exe") Then
        MsgBox("Impossibile avviare il programma d'installazione aggiornamenti (Busagg.exe) in quanto non possiede firma digitale valida.", vbExclamation)
        Return True
      End If
      'davide: non serve
      'If Not System.IO.File.Exists(strBusAggdir & "Busagm.exe") Then
      '  MsgBox("Attenzione: non è possibile avviare la procedura di aggiornamento in quanto non è stato trovato il file '" & strBusAggdir & "Busagm.exe'", vbExclamation)
      '  Return True
      'End If

      '----------------------------------
      'lancio Busagg.exe e chiudo Busnet.exe
      MsgBox("Attenzione: sul server è presente un aggiornamento di Business. Business CUBE verrà chiuso e verrà avviata la procedura di aggiornamento", MsgBoxStyle.Information)

      'vecchio sistema: quello precedente dava errrore quando da xp utente di tipo user devo sfogliare la rete per accedere ad aggnumber.txt, quando in realtà tutto l'ambiente è impostato correttamente
      'Dim a As New ProcessStartInfo()
      'Dim p As New Process()
      'Try
      '  a.FileName = strBusAggdir & "Busagg.exe"
      '  a.Arguments = strProfilo & " /B"
      '  a.WorkingDirectory = strBusDir
      '  a.LoadUserProfile = True
      '  a.UseShellExecute = False
      '  a.RedirectStandardOutput = True
      '  If strAggUser.Trim <> "" Then
      '    i = strAggUser.IndexOf("\")
      '    If i < 0 Then
      '      strDomain = ""
      '      strUser = strAggUser
      '    Else
      '      strDomain = strAggUser.Substring(0, i)
      '      strUser = strAggUser.Substring(i + 1)
      '    End If
      '    a.Domain = strDomain
      '    a.UserName = strUser

      '    PwdDeCript(strAggPwd, strAggPwd)

      '    For i = 0 To strAggPwd.Length - 1
      '      s.AppendChar(strAggPwd(i))
      '    Next
      '    a.Password = s.Copy

      '    If System.Environment.OSVersion.Version.Major >= 6 Then
      '      p.StartInfo = a      'in XP così da errore !!!! handle non valido
      '      p.Start()
      '    Else
      '      'specifico per windows XP
      '      Process.Start(strBusAggdir & "Busagg.exe", strProfilo & " /B", strUser, s.Copy, strDomain)
      '    End If
      '  Else
      '    If System.Environment.OSVersion.Version.Major >= 6 Then
      '      p.StartInfo = a      'in XP così da errore !!!! handle non valido
      '      p.Start()
      '    Else
      '      'specifico per windows XP: non posso fare il test per vista, visto che se ho lo uacattivo non sono mai un admin
      '      'se l'utente passatomi non è una admin do un messaggio e non faccio partire l'aggiornamento
      '      Dim user As New System.Security.Principal.WindowsPrincipal(System.Security.Principal.WindowsIdentity.GetCurrent())
      '      If user.IsInRole(System.Security.Principal.WindowsBuiltInRole.Administrator) = False Then
      '        MsgBox("Attenzione: dovrebbero essere caricati degli aggiornamenti di Business NET" & vbCrLf & _
      '               "ma l'utente di windows corrente non è un amministratore" & vbCrLf & _
      '               "e nelle proprietà degli aggiornamenti non è stato indicato" & vbCrLf & _
      '               "un utente di tipo administrator. Gli aggiornamenti non verranno caricati", MsgBoxStyle.Critical)
      '        Return True
      '      Else
      '        Process.Start(strBusAggdir & "Busagg.exe", strProfilo & " /B")
      '      End If
      '    End If
      '  End If    'If strAggUser.Trim <> "" Then

      '  Me.Close()      'chiudo: l'utente dovrà riavviare bus a mano dopo la fine degli aggiornamenti
      '  Return False    'in questo modo busnet viene chiuso...
      'Catch ex As Exception
      '  MsgBox("Errore in fase di avvio '" & strBusAggdir & "Busagg.exe" & "' per installazione aggiornamenti." & vbCrLf & "Gli aggiornamenti non verranno caricati." & vbCrLf & vbCrLf & vbCrLf & "Errore: " & vbCrLf & ex.ToString)
      '  Return True
      'End Try

      'nuovo sistema
      Dim strParam As String = "/FILEPROFILO=""" & Application.StartupPath & "\" & strProfilo & ".cfg" & """ /BUSNET /SILENT"
      Try
        Dim pSecurityAttributes As SECURITY_ATTRIBUTES
        pSecurityAttributes.bInheritHandle = True
        pSecurityAttributes.nLength = Marshal.SizeOf(pSecurityAttributes)
        pSecurityAttributes.lpSecurityDescriptor = IntPtr.Zero
        Dim pStartInfo As STARTUP_INFO = Nothing
        pStartInfo.cb = Len(pStartInfo)
        Dim pProcessInfo As PROCESS_INFO

        If strAggUser.Trim <> "" Then
          i = strAggUser.IndexOf("\")
          If i < 0 Then
            strDomain = ""
            strUser = strAggUser
          Else
            strDomain = strAggUser.Substring(0, i)
            strUser = strAggUser.Substring(i + 1)
          End If

          oStd.PwdDeCript(strAggPwd, strAggPwd)


          'avvio gli aggiornamenti con credenziali specificate
          Dim iReturn As Integer = CreateProcessWithLogonW(strUser, strDomain, strAggPwd,
                                          LOGON_WITH_PROFILE, Nothing,
                                          strBusAggdir & "Busagg.exe " & strParam,
                                          CREATE_DEFAULT_ERROR_MODE, Nothing,
                                          strBusAggdir, pStartInfo, pProcessInfo)
          If iReturn = 0 Then
            Dim lLastError As Long = GetLastError()
            Select Case lLastError
              Case 1326, 1783
                MsgBox("Errore in fase di avvio aggiornamenti:" & vbCrLf & "dominio, utente o password errati o non impostati (o utente non di tipo administrator). Errore: " & vbCrLf & GetLastError().ToString, vbExclamation, "Business CUBE")

              Case 2
                MsgBox("Errore in fase di avvio aggiornamenti:" & vbCrLf & " programma " & strBusAggdir & "Busagg.exe' non trovato. Errore: " & vbCrLf & GetLastError().ToString, vbExclamation, "Business CUBE")

              Case Else
                Dim strMessaggio As String = Space(200)
                FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0&, lLastError, 0&, strMessaggio, 200, 0&)

                MsgBox("Errore in fase di avvio aggiornamenti:" & vbCrLf & " Errore: " & vbCrLf & lLastError.ToString & " - Messaggio:" & strMessaggio.Trim, vbExclamation, "Business CUBE")
            End Select
            Return True 'proseguo senza avviare gli aggiornamenti
          End If
        Else
          'non ho indicato credenziali: se l'utente non è un administrator verrà generato un errore non gestito
          Process.Start(strBusAggdir & "Busagg.exe", strParam)
        End If    'If strAggUser.Trim <> "" Then

        Me.Close()      'chiudo: l'utente dovrà riavviare bus a mano dopo la fine degli aggiornamenti
        Return False    'in questo modo busnet viene chiuso...

      Catch ex As Exception
        MsgBox("Errore in fase di avvio '" & strBusAggdir & "Busagg.exe" & "' per installazione aggiornamenti." & vbCrLf & "Gli aggiornamenti non verranno caricati." & vbCrLf & vbCrLf & vbCrLf & "Errore: " & vbCrLf & ex.ToString, vbCritical, "Business CUBE")
        Return True
      End Try

    Catch ex As Exception
      MsgBox("Error starting Business NET - CaricaAgg: " & ex.ToString, vbCritical, "Business CUBE")
      Me.Close()
    End Try
  End Function

  Public Function CheckBusProfilo(ByVal strComando As String, ByRef bUserXXX As Boolean) As String
    'scompone la riga di comando in : opnome passwd codazi profilo
    Dim nPos As Integer
    Dim nPos1 As Integer
    Dim nPos2 As Integer
    Dim nPos3 As Integer
    Dim nPos4 As Integer
    Dim strC As String

    CheckBusProfilo = "Business"

    If Len(Trim(strComando)) = 0 Then Exit Function

    'Legge la linea di comando
    strC = strComando
    nPos = InStr(1, strC, " ")
    If nPos <> 0 Then
      nPos1 = InStr(nPos + 1, strC, " ")
      If Trim(Microsoft.VisualBasic.Left(strC, nPos - 1)).ToUpper = "XXX" Then bUserXXX = True
      'gstrOpnome = Trim(Left$(strC, nPos - 1))
    End If
    If nPos1 <> 0 Then
      nPos2 = InStr(nPos1 + 1, strC, " ")
      'gstrPasswd = Mid$(strC, nPos + 1, nPos1 - nPos - 1)
    End If
    If nPos2 <> 0 Then
      nPos3 = InStr(nPos2 + 1, strC, " ")
      'gstrCodazi = Mid$(strC, nPos1 + 1, nPos2 - nPos1 - 1)
    Else
      'gstrCodazi = Mid$(strC, nPos1 + 1)
    End If
    'Aggiunto : profilo
    If nPos3 <> 0 Then
      nPos4 = InStr(nPos3 + 1, strC, " ")
      CheckBusProfilo = Mid$(strC, nPos2 + 1, nPos3 - nPos2 - 1)
    Else
      If nPos2 <> 0 Then
        CheckBusProfilo = Mid$(strC, nPos2 + 1)
      End If
    End If
    'Aggiunto: programma
    'Aggiunto: restricted /R puo' esserci solo se c'e' programma,
    'Aggiunto: parametro/i con cui chiamare il programma, può esserci solo se c'è programma
    'NOTA:     se c'è restricted il/i parametro/i devono essere scritti dopo
    'Esempio: Roby . Prova business bs--paga /R Apri;3
    '  If nPos4 <> 0 Then
    '    gstrProgramma = Mid$(strC, nPos3 + 1, nPos4 - nPos3 - 1)
    '    nPos5 = InStr(nPos4 + 1, strC, " ")
    '    If nPos5 <> 0 Then
    '      gstrRestr = Mid$(strC, nPos4 + 1, nPos5 - nPos4 - 1)
    '      gstrParam = Mid$(strC, nPos5 + 1)
    '    Else
    '      strAppo = Mid$(strC, nPos4 + 1)
    '      If strAppo = "/R" Then
    '        gstrRestr = Mid$(strC, nPos4 + 1)
    '      Else
    '        gstrParam = Mid$(strC, nPos4 + 1)
    '      End If
    '    End If
    '  Else
    '    If nPos3 <> 0 Then
    '     gstrProgramma = Mid$(strC, nPos3 + 1)
    '    End If
    '  End If
    'Ogni campo con . viene considerato poi stringa vuota ...
    '  If gstrPasswd = "." Then gstrPasswd = ""
    '  If gstrOpnome = "." Then gstrOpnome = ""
    '  If gstrCodazi = "." Then gstrCodazi = ""
    If CheckBusProfilo = "." Or CheckBusProfilo = "" Then CheckBusProfilo = "Business"
    '  If gstrProgramma = "." Then gstrProgramma = ""
    '  If gstrRestr = "." Then gstrRestr = ""
    '  If gstrParam = "." Then gstrParam = ""
  End Function

#Region "Mutex"
  Public Function NtsMutexCrea(ByVal strMutexName As String, ByRef bExist As Boolean) As Boolean
    Try
      NtsMutexCrea = False
      bExist = False
      If oStd.IsDebugMode() Then Return False

      'Crea il mutex Global\
      Dim bNew As Boolean = False
      oMutexG = New System.Threading.Mutex(True, "Global\" & strMutexName, bNew)
      If bNew = False Then
        bExist = True
      End If

      'Crea il mutex Local\
      oMutexL = New System.Threading.Mutex(False, strMutexName)

    Catch ex As Exception
      NtsMutexCrea = False
      bExist = False
    End Try
  End Function
  Public Function NtsMutexElimina() As Boolean
    NtsMutexElimina = False
    Try

      'Rilascia/chiude il mutex (aperto con CreateMutex) \Global
      If Not oMutexG Is Nothing Then
        oMutexG.Close()
        oMutexG.Dispose()
        oMutexG = Nothing
      End If

      'Rilascia/chiude il mutex (aperto con CreateMutex) \Local
      If Not oMutexL Is Nothing Then
        oMutexL.Close()
        oMutexL.Dispose()
        oMutexL = Nothing
      End If

      Return True
    Catch ex As Exception
      'Non segnala nessun errore ed esce come se NON avesse trovato il mutex da eliminare
      'Tanto viene comunque eliminato dal sistema operativo quando il processo viene chiuso
      'MsgBox "Errore nella routine NtsMutexElimina: " & Err.Number & " - " & Err.Description & "(LastDllError " & Err.LastDllError & ")", vbCritical, "Messaggio di Business"
      Return False
    End Try
  End Function
  Public Function NtsMutexTrova(ByVal strMutexName As String) As Boolean
    Try
      'Controlla il mutex Local\
      Dim bNew As Boolean = True
      Dim oTmp As New System.Threading.Mutex(True, strMutexName, bNew)
      If Not oTmp Is Nothing Then
        oTmp.Close()
        oTmp.Dispose()
        oTmp = Nothing
      End If

      If bNew = False Then Return True

      oTmp = New System.Threading.Mutex(True, "Global\" & strMutexName, bNew)
      If Not oTmp Is Nothing Then
        oTmp.Close()
        oTmp.Dispose()
        oTmp = Nothing
      End If

      If bNew = False Then Return True

      Return False
    Catch ex As Exception
      'Non segnala nessun errore ed esce come se NON avesse trovato il mutex
      'MsgBox "Errore nella routine TrovaNtsMutex: " & Err.Number & " - " & Err.Description & "(LastDllErro " & Err.LastDllError & ")", vbCritical, "Messaggio di Business"
      Return False
    End Try
  End Function
#End Region

  Public Sub WriteMsgBoxToLog(ByVal strMsg As String, ByVal bAccoda As Boolean)
    'se sono stato avviato in modalità batch (tipo busbatch) i messaggi vengono a finire tutti qui (eccetto quelli du busnet.exe)
    Dim lw1 As New StreamWriter(LogDirP & "\BusNetBatch_" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString & ".log", bAccoda)
    lw1.WriteLine(DateTime.Now.ToString & " " & strMsg)
    lw1.Flush()
    lw1.Close()
  End Sub

  Public Function VerifyFileSignature(ByVal filePath As String) As Boolean
    Try
      Dim signedFile As New X509Certificate2(filePath)
      Return signedFile.SubjectName.Name.Contains("CN=NTS Informatica Srl") OrElse signedFile.SubjectName.Name.Contains("CN=CNA SERVIZI SCRL")
    Catch ex As Exception
      Return False
    End Try
  End Function
End Class
