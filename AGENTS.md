# Istruzioni progetto Buoninfante

## Contesto

Questa repository contiene la personalizzazione Business CUBE Buoninfante.

- Release Business: SR8 CU4.
- Solution: `Buoninfante.sln`.
- Ambiente: Visual Studio 2019, .NET Framework 4.8.
- Configurazione concordata: `Debug|Any CPU`.
- Avviatore locale: progetto `BUSCUBE`.
- Skill da utilizzare per attività Business: `business-cube`.

## Riferimenti autorevoli

Prima di analizzare o modificare il progetto, leggere il profilo:

`C:\Users\Utente\source\Assistente_Sviluppo\Business\projects\BUONINFANTE.md`

Per firme, comportamento e compatibilità, confrontare sempre i sorgenti standard
SR8 CU4 censiti in:

`C:\Users\Utente\source\Assistente_Sviluppo\Business\knowledge\releases\SR8-CU4\SOURCES.md`

Consultare inoltre il rapporto di analisi quando la richiesta riguarda rischi,
criticità note o riuso di pattern:

`C:\Users\Utente\source\Assistente_Sviluppo\Business\reports\BUONINFANTE-SR8-CU4-ANALISI.md`

## Regole operative

- Le richieste esplicite dell'utente prevalgono sulle indicazioni generali della
  skill; segnalare eventuali conflitti prima di cambiare direzione.
- Non modificare i sorgenti standard Business.
- Non copiare firme o pattern da una CU diversa senza confronto con SR8 CU4.
- Conservare la separazione tra Form, Entity e DAL.
- Prima di modificare form secondarie, verificare il mapping effettivo in
  `Dllmap.ini` e gli eventuali XML esportati dall'editor NTS.
- Non inventare identificativi `oApp.Tr`.
- Non assumere che il codice esistente sia uno standard approvato: distinguere
  sempre codice osservato, regola confermata e ipotesi.
- Preservare le modifiche locali già presenti e limitare ogni intervento ai file
  necessari per la richiesta.

## Verifica

- Per modifiche al codice, compilare almeno `Buoninfante.sln` in
  `Debug|Any CPU`, salvo diversa indicazione dell'utente.
- Eseguire i test pertinenti disponibili e dichiarare chiaramente ciò che non è
  stato possibile verificare nell'interfaccia o nel database.
- Non usare ambienti o dati di produzione per prove distruttive.
