# PEC IMAP Backup PySide6

Utility desktop in Python/PySide6 per fare backup locale di caselle IMAP/PEC.

## Funzioni principali

- Salvataggio dei messaggi in formato `.eml` RAW, byte-per-byte, come “Salva con nome” di Thunderbird.
- Conservazione completa di testo, allegati, MIME, firme PEC e messaggi allegati.
- Filtro per data: salva i messaggi prima di una data scelta.
- Filtro opzionale per dimensione: salva solo messaggi più grandi di una soglia in B/KB/MB/GB.
- Selezione delle cartelle IMAP reali tramite lista spuntabile caricata dal server.
- Deduplica con SHA-256, utile se il backup si interrompe e viene rilanciato.
- Indice Excel `indice_backup_pec.xlsx` con metadati, percorso locale, SHA-256 e `mittente_vero`.
- Parsing PEC da `daticert.xml` / `postacert.eml` per evitare mittenti generici come `posta-certificata@...`.
- Anteprima tabellare e doppio click su una riga per vedere la sorgente RAW del messaggio.
- Pulsante Stop con interruzione cooperativa.
- Log ridimensionabile e cancellabile.
- Selettore lingua Italiano / English con file JSON in `i18n/`.
- Icona applicazione in `resources/`.

## Installazione

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python pec_imap_backup_pyside6.py
```

Su Windows:

```bat
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python pec_imap_backup_pyside6.py
```

## Uso consigliato

1. Inserisci host, porta, sicurezza, utente e password.
2. Scegli la lingua dall'apposito selettore.
3. Premi **Test / lista cartelle** per caricare i nomi reali delle mailbox IMAP.
4. Spunta le cartelle da processare.
5. Imposta la data “Salva messaggi prima del”.
6. Attiva il filtro dimensione se serve.
7. Premi **Anteprima** per vedere cosa verrà lavorato.
8. Premi **Esegui backup**.
9. Solo se il backup termina senza errori, il programma propone la marcatura `\Deleted`.
10. `EXPUNGE` resta separato e va abilitato esplicitamente.

## Lingue

Le traduzioni sono file JSON:

- `i18n/it.json`
- `i18n/en.json`

Per correggere o estendere una traduzione basta modificare il valore associato alla chiave. Le chiavi devono restare identiche.

## Note sulla cancellazione IMAP

Il backup e la cancellazione sono fasi separate. Il programma non marca nulla come `\Deleted` durante il salvataggio. La marcatura viene proposta solo alla fine, dopo backup e verifica, e richiede conferma esplicita.

`EXPUNGE` elimina definitivamente dal server i messaggi già marcati `\Deleted`. Per sicurezza è disattivato di default.

## Debug e crash Qt/PySide6

Sono disponibili opzioni di avvio:

```bash
python pec_imap_backup_pyside6.py --safe-mode
python pec_imap_backup_pyside6.py --reset-settings
python pec_imap_backup_pyside6.py --qt-debug-plugins
python pec_imap_backup_pyside6.py --force-xcb
python pec_imap_backup_pyside6.py --force-wayland
```

Il log crash viene scritto in `pec_imap_backup_crash.log`, oppure nel percorso indicato dalla variabile ambiente:

```bash
PEC_IMAP_BACKUP_CRASH_LOG=/percorso/log.txt python pec_imap_backup_pyside6.py
```

## Sicurezza password

Se abiliti “salva password”, la password viene memorizzata localmente tramite QSettings con offuscazione base64/XOR. Non è cifratura forte. Per maggiore sicurezza, lascia disattivata l'opzione e inserisci la password a mano.

## Nota codice

I commenti nel codice sono in inglese. I messaggi utente passano dal sistema di traduzione JSON.
