# Magazzino Elettronica — ITIS Cannizzaro Colleferro

Gestionale web per l'inventario dei laboratori di elettronica. Importa automaticamente i 725 componenti dal file Excel fornito al primo avvio e usa un database SQLite locale.

## Funzioni

- inventario ricercabile con giacenze e ubicazioni;
- prelievi, restituzioni, carichi e rettifiche associati all'operatore;
- storico movimenti protetto da modifica e cancellazione anche a livello database;
- soglia minima e quantità obiettivo configurabili per ogni componente;
- dashboard per esauriti, sotto soglia e configurazioni incomplete;
- distinta di riordino calcolata fino alla quantità obiettivo;
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

Un articolo compare nella distinta quando la giacenza è minore o uguale alla soglia e la quantità obiettivo è superiore alla giacenza. La quantità proposta è:

`quantità obiettivo - giacenza attuale`

## Uso nella rete scolastica

L'app ascolta su tutte le interfacce alla porta 3000. Per un uso stabile da più computer va eseguita su un PC/server della scuola, protetta tramite HTTPS e con una chiave persistente impostata in `MAGAZZINO_SECRET_KEY`. Il file SQLite deve restare sul disco locale del server, non su una cartella di rete sincronizzata.
