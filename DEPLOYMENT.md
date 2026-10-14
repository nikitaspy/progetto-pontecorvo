# Pubblicazione sul server scolastico

Il gestionale è un'applicazione Python con database e non può essere copiato dentro una normale pagina HTML, WordPress o uno spazio FTP statico. Il sito istituzionale può aggiungere un collegamento, ad esempio **Magazzino elettronica**, verso un sottodominio dedicato:

`https://magazzino.itis-cannizzaro.it`

Sul server devono essere eseguiti l'applicazione e Waitress, mantenendo la cartella `instance` su disco persistente. Un reverse proxy come IIS, Caddy o Nginx gestisce il certificato HTTPS e inoltra le richieste all'applicazione. La porta interna 3000 non va esposta direttamente su Internet. Il progetto deve trovarsi in una cartella applicativa riservata: sul sito pubblico compare soltanto il collegamento al gestionale.

## Configurazione consigliata

- Un solo server o macchina virtuale sempre accesa.
- Python 3.11 o successivo e dipendenze di `requirements.txt`.
- Waitress avviato come servizio di sistema, con riavvio automatico.
- IIS, Caddy o Nginx davanti a Waitress, con HTTPS sulla porta 443.
- Nome DNS stabile, preferibilmente un sottodominio della scuola.
- Database `instance/magazzino.db` sul disco locale del server.
- Backup giornaliero su una destinazione separata e prova periodica di ripristino.

La cartella `instance` contiene il database e la chiave usata per le sessioni. Per una semplice copia manuale, arrestare prima il servizio e copiare l'intera cartella. Per backup a caldo occorre invece usare l'API di backup di SQLite o uno strumento che la supporti; la sola copia di `magazzino.db` mentre il servizio è attivo può omettere dati ancora presenti nel file WAL.

Le variabili principali sono:

```text
MAGAZZINO_SECRET_KEY=<chiave casuale lunga e permanente>
MAGAZZINO_PUBLIC_URL=https://magazzino.itis-cannizzaro.it
MAGAZZINO_HTTPS=1
MAGAZZINO_TRUSTED_HOSTS=magazzino.itis-cannizzaro.it
MAGAZZINO_HOST=127.0.0.1
MAGAZZINO_THREADS=8
PORT=3000
```

Queste sono variabili dell'ambiente del processo o del servizio: il progetto non legge automaticamente un file `.env`. Se l'indirizzo pubblico viene salvato anche dalla pagina **Etichette QR**, il valore memorizzato nel database ha la precedenza su `MAGAZZINO_PUBLIC_URL`. Prima di stampare le etichette, verificare quindi che la pagina mostri il dominio definitivo.

Avvio applicativo:

```text
python serve.py
```

Al primo avvio lasciare il servizio raggiungibile soltanto dal server o dalla rete amministrativa, aprire `/setup` e creare subito l'amministratore iniziale. Solo dopo questa configurazione si abilita il dominio pubblico nel reverse proxy. Sul proxy è inoltre opportuno limitare i tentativi ripetuti verso `/login`.

Su Windows il servizio può essere registrato con WinSW, NSSM o l'Utilità di pianificazione; su Linux si usa normalmente `systemd`. Configurare sempre come directory di lavoro la cartella del progetto e avviare l'interprete dell'ambiente virtuale con `serve.py`. Il servizio elimina la finestra dei comandi, parte con il server e continua a funzionare anche senza un utente collegato. Per una prova su un singolo PC Windows è disponibile `avvia-windows-silenzioso.vbs`, ma per il servizio definitivo va coinvolto il referente informatico della scuola.

## Accesso dalla rete Wi-Fi

Per l'uso solo interno, i telefoni devono poter raggiungere il server dalla Wi-Fi docenti. L'ufficio tecnico deve verificare che le VLAN e l'isolamento client consentano il traffico dai dispositivi al server sulla porta HTTPS 443. `localhost` sul telefono indica il telefono stesso e non il computer del magazzino.

SQLite è adeguato per un singolo server e un numero normale di docenti. Il gestionale usa transazioni brevi, attesa sui blocchi e modalità WAL. Il database non deve essere collocato su OneDrive, SharePoint, NAS/SMB o altre cartelle sincronizzate. Se in futuro serviranno più server applicativi, il database andrà trasferito a PostgreSQL.

## Flusso QR

1. Avviare il gestionale in accesso riservato e creare l'amministratore iniziale.
2. Pubblicarlo tramite HTTPS con l'indirizzo definitivo.
3. Accedere come amministratore e creare un account personale `Docente / operatore` per ogni docente.
4. Aprire **Etichette QR** e salvare l'indirizzo HTTPS definitivo.
5. Filtrare e selezionare i componenti.
6. Scaricare e stampare il PDF delle etichette.

Il QR contiene soltanto l'indirizzo stabile del componente. Giacenza, posizione e caratteristiche vengono lette dal database al momento della scansione e rimangono quindi aggiornate. La scheda del componente è pubblica e può essere consultata da chiunque disponga del QR o del collegamento. Il login è obbligatorio per modificare la giacenza con un prelievo o una restituzione; ogni movimento resta associato all'account autenticato. Dopo l'accesso, il menu mostra le funzioni consentite dal ruolo e permette anche di registrare più componenti nella stessa operazione. Soltanto gli amministratori possono creare, modificare, eliminare o cambiare i privilegi degli utenti; gli account già presenti nello storico vengono disattivati invece di essere rimossi.
