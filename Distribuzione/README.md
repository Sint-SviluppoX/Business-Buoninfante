# DLL personalizzazioni Buoninfante

Questa cartella contiene DLL e PDB dei componenti custom inclusi in
`Buoninfante.sln`.

La compilazione di un progetto di tipo `Library` aggiorna automaticamente i
relativi file tramite `Directory.Build.targets`. Prima di creare un commit che
modifica una DLL, compilare la solution in configurazione `Debug|Any CPU` e
includere anche gli artefatti aggiornati di questa cartella.

`BUSCUBE` non viene copiato perché è l'eseguibile di avvio, non un componente
DLL personalizzato.
