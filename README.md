# Magazzino Elettronica — ITIS Cannizzaro Colleferro

Versione corrente: **1.3.0**.

Gestionale web per l'inventario dei laboratori di elettronica. Importa automaticamente i 725 componenti dal file Excel fornito al primo avvio e usa un database SQLite locale.

## Funzioni

- inventario ricercabile con giacenze e ubicazioni;
- prelievi, restituzioni, carichi e rettifiche associati all'operatore;
- storico movimenti protetto da modifica e cancellazione anche a livello database;
- soglia minima e quantità obiettivo configurabili per ogni componente;
- dashboard per esauriti, sotto soglia e configurazioni incomplete;
- distinta da completare con tutti i componenti che richiedono un riordino e non hanno ancora una quantità confermata;
- selezione multipla o totale della distinta con una quantità desiderata comune, mantenendo l'inserimento per singolo componente;
- modulo di riordino separato con le sole quantità confermate, modificabile e stampabile in HTML, Excel e PDF;
- esportazione della distinta in Excel e PDF;
- account individuali con ruoli amministratore e operatore.

## Avvio

Richiede Python 3.11 o successivo.

Su Windows, estrarre prima l'intero archivio ZIP in una cartella. Dalla cartella estratta è possibile fare doppio clic su `avvia-windows.bat`: al primo avvio prepara l'ambiente, installa le dipendenze e apre il browser. Non eseguire il file direttamente dall'anteprima dell'archivio ZIP di Windows.

```bash
python -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/python app.py
```

Aprire `http://localhost:3000`. Al primo accesso l'app chiede di creare l'amministratore. Non esistono credenziali predefinite.

Il database viene creato in `instance/magazzino.db`. Per un backup completo è sufficiente copiare questo file quando l'applicazione non sta registrando operazioni.

## Regole di riordino

L'importazione imposta la quantità obiettivo uguale alla giacenza originale dell'Excel. La soglia iniziale viene letta dalla colonna `SOGLIA ALLARME`, oppure impostata a zero quando è vuota. Gli articoli originariamente esauriti restano con obiettivo zero e sono evidenziati nella dashboard per una configurazione manuale.

Un articolo compare sempre nella distinta quando la giacenza è zero. Compare inoltre quando la giacenza è minore o uguale alla soglia. Se l'obiettivo è configurato, la distinta propone una quantità; altrimenti la lascia da definire senza inventare un valore.

Quando l'obiettivo è configurato, la quantità proposta è:

`quantità obiettivo - giacenza attuale`

La quantità proposta è solo un suggerimento. Quando l'amministratore conferma una quantità, il componente passa dalla **Distinta da completare** al **Modulo riordino**. La quantità desiderata viene memorizzata separatamente dalla soglia e dall'obiettivo di scorta. Nel modulo può essere modificata, stampata o esportata; con **Rimetti in distinta** la conferma viene rimossa e il componente torna tra quelli da completare.

## Aggiornamento dalla versione 1.2

Per conservare utenti, giacenze e movimenti, copiare la cartella `instance` della versione precedente nella nuova cartella del progetto prima dell'avvio. Le quantità d'ordine della versione 1.2 che possono essere riconosciute dai dati e dallo storico vengono trasferite automaticamente nel nuovo modulo. Poiché quella versione memorizzava nello stesso campo l'obiettivo di scorta e la quantità confermata, gli eventuali casi non distinguibili restano nella distinta e vanno confermati di nuovo.

## Uso nella rete scolastica

L'app ascolta su tutte le interfacce alla porta 3000. Per un uso stabile da più computer va eseguita su un PC/server della scuola, protetta tramite HTTPS e con una chiave persistente impostata in `MAGAZZINO_SECRET_KEY`. Il file SQLite deve restare sul disco locale del server, non su una cartella di rete sincronizzata.
