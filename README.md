# Magazzino Elettronica — ITIS Cannizzaro Colleferro

Versione corrente: **1.4.0**.

Gestionale web per l'inventario dei laboratori di elettronica. Importa automaticamente i 725 componenti dal file Excel fornito al primo avvio e usa un database SQLite locale.

## Funzioni

- inventario ricercabile con giacenze e ubicazioni;
- prelievi, restituzioni, carichi e rettifiche associati all'operatore;
- prelievi e riconsegne multipli, con più pezzi uguali o componenti diversi registrati integralmente oppure annullati in caso di errore;
- storico movimenti protetto da modifica e cancellazione anche a livello database;
- soglia minima e quantità obiettivo configurabili per ogni componente;
- dashboard per esauriti, sotto soglia e configurazioni incomplete;
- distinta da completare con tutti i componenti che richiedono un riordino e non hanno ancora una quantità confermata;
- selezione multipla o totale della distinta con una quantità desiderata comune, mantenendo l'inserimento per singolo componente;
- modulo di riordino separato con le sole quantità confermate, modificabile e stampabile in HTML, Excel e PDF;
- esportazione della distinta in Excel e PDF;
- account individuali con ruoli amministratore e operatore, gestibili soltanto dagli amministratori;
- scheda pubblica del componente tramite QR e prelievo o restituzione da smartphone dopo il login;
- generazione PDF delle etichette QR singole o multiple;
- protezione dai doppi invii e permessi docenti limitati a prelievi e restituzioni;
- server Waitress e avvio Windows silenzioso senza finestra dei comandi.

## Avvio

Richiede Python 3.11 o successivo.

Su Windows, estrarre prima l'intero archivio ZIP in una cartella locale stabile, ad esempio `C:\MagazzinoTPS`. Dalla cartella estratta è possibile fare doppio clic su `avvia-windows.bat`: al primo avvio prepara l'ambiente, installa le dipendenze e apre il browser. Non eseguire il file direttamente dall'anteprima dell'archivio ZIP di Windows e non collocare il progetto in una cartella sincronizzata da OneDrive.

Dopo il primo avvio riuscito si può usare `avvia-windows-silenzioso.vbs`: il gestionale parte in background e il browser si apre senza mostrare la finestra dei comandi. Per arrestare questa esecuzione usare `ferma-windows.bat`, in particolare prima di aggiornare il programma o copiarne il backup. Se l'avvio silenzioso non riesce, eseguire il file `.bat` normale per leggere l'errore; il log è disponibile in `instance/avvio.log`.

```bash
python -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/python serve.py
```

Aprire `http://localhost:3000`. Al primo accesso l'app chiede di creare l'amministratore. Non esistono credenziali predefinite.

Il database viene creato in `instance/magazzino.db`. Per un backup manuale sicuro, arrestare l'applicazione e copiare l'intera cartella `instance`; in questo modo vengono conservati anche la chiave delle sessioni e gli eventuali file SQLite temporanei.

## Smartphone e QR

L'amministratore configura l'indirizzo raggiungibile dai telefoni nella pagina **Etichette QR**, seleziona i componenti e scarica il PDF da stampare. Ogni QR contiene un collegamento stabile come `https://magazzino.scuola.it/q/TPS0070`. Chiunque inquadri il codice può consultare la scheda pubblica con descrizione, valore, ubicazione e giacenza aggiornata. Per registrare un prelievo o una restituzione occorre accedere con un account personale; il movimento viene associato automaticamente al docente. Dopo l'accesso sono disponibili anche il menu e le pagine consentite dal proprio ruolo.

Gli amministratori possono creare, modificare, disattivare o eliminare gli account e assegnare i privilegi. Quando un account possiede già movimenti, l'eliminazione lo disattiva e conserva lo storico. Gli operatori possono consultare il magazzino e registrare prelievi e restituzioni, anche per più componenti nella stessa operazione, ma non possono amministrare gli utenti né eseguire carichi o rettifiche.

L'indirizzo `localhost` funziona soltanto sul computer che esegue il gestionale e non deve essere stampato nei QR. Per usare gli smartphone serve un server raggiungibile dalla rete Wi-Fi scolastica oppure un indirizzo pubblico HTTPS.

## Regole di riordino

L'importazione imposta la quantità obiettivo uguale alla giacenza originale dell'Excel. La soglia iniziale viene letta dalla colonna `SOGLIA ALLARME`, oppure impostata a zero quando è vuota. Gli articoli originariamente esauriti restano con obiettivo zero e sono evidenziati nella dashboard per una configurazione manuale.

Un articolo compare sempre nella distinta quando la giacenza è zero. Compare inoltre quando la giacenza è minore o uguale alla soglia. Se l'obiettivo è configurato, la distinta propone una quantità; altrimenti la lascia da definire senza inventare un valore.

Quando l'obiettivo è configurato, la quantità proposta è:

`quantità obiettivo - giacenza attuale`

La quantità proposta è solo un suggerimento. Quando l'amministratore conferma una quantità, il componente passa dalla **Distinta da completare** al **Modulo riordino**. La quantità desiderata viene memorizzata separatamente dalla soglia e dall'obiettivo di scorta. Nel modulo può essere modificata, stampata o esportata; con **Rimetti in distinta** la conferma viene rimossa e il componente torna tra quelli da completare.

## Aggiornamento da una versione precedente

Arrestare la versione precedente e copiare la sua cartella `instance` nella nuova cartella del progetto prima dell'avvio, così da conservare utenti, giacenze e movimenti. Le quantità d'ordine della versione 1.2 che possono essere riconosciute dai dati e dallo storico vengono trasferite automaticamente nel nuovo modulo. Poiché quella versione memorizzava nello stesso campo l'obiettivo di scorta e la quantità confermata, gli eventuali casi non distinguibili restano nella distinta e vanno confermati di nuovo.

## Uso nella rete scolastica

Per impostazione predefinita `serve.py` ascolta su tutte le interfacce alla porta 3000. Per un uso stabile da più computer va eseguito su un PC/server della scuola, protetto tramite HTTPS e con una chiave persistente impostata in `MAGAZZINO_SECRET_KEY`. Dietro un reverse proxy, l'ascolto applicativo va limitato a `127.0.0.1`. Il file SQLite deve restare sul disco locale del server, non su una cartella di rete sincronizzata.

Le indicazioni complete per la pubblicazione sono in [DEPLOYMENT.md](DEPLOYMENT.md).
