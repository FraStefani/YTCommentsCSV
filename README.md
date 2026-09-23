# YTCommentsCSV

Scarica **tutti** i commenti di un video YouTube e li salva in un report Excel colorato:
una tabella con i dati del video in alto, la tabella dei commenti sotto.

Il nome del file è `<nome canale>_<data attuale>.xlsx`, ad esempio
`Marco Montemagno_2026-09-23.xlsx`.

> **Perché .xlsx e non .csv**: un file `.csv` è testo puro, non può contenere colori né due
> tabelle distinte. I colori richiesti sono possibili solo in un file Excel.

> Se cerchi le istruzioni pratiche passo a passo ("apro il terminale, incollo il link,
> ottengo il file"), sono in [COME_USARE.md](COME_USARE.md).

## Requisiti

- Python 3.10 o superiore (testato su 3.14)
- `openpyxl` — l'unica dipendenza:
  ```
  pip install -r requirements.txt
  ```

## API key (gratuita, ~2 minuti)

1. Apri <https://console.cloud.google.com/> e crea un progetto.
2. **API e servizi → Libreria →** cerca *YouTube Data API v3* → **Abilita**.
3. **API e servizi → Credenziali → Crea credenziali → Chiave API**.
4. Copia `.env.example` in `.env` e incolla la chiave:
   ```
   YOUTUBE_API_KEY=AIza...
   ```

In alternativa: variabile d'ambiente `YOUTUBE_API_KEY` oppure opzione `--api-key`.
Il file `.env` è escluso da git e la chiave non viene mai scritta nel report né nei messaggi
di errore.

**Quota**: 10.000 unità al giorno, 1 unità per chiamata, 100 commenti per chiamata → circa
un milione di commenti al giorno. La quota si azzera a mezzanotte, ora del Pacifico.

## Uso

```bash
python ytcomments.py "https://youtu.be/ByZrTRph2-4"
python ytcomments.py "https://youtu.be/ByZrTRph2-4" --max 50      # prova rapida
python ytcomments.py ByZrTRph2-4 --no-replies --out C:\report
```

| Opzione | Effetto |
|---|---|
| `--max N` | si ferma dopo N commenti (default: tutti) |
| `--no-replies` | solo commenti di primo livello, senza risposte |
| `--order time\|relevance` | ordine di scaricamento dall'API (default `time`) |
| `--out DIR` | cartella di destinazione (default: quella dello script) |
| `--name NOME` | nome file senza estensione, al posto di `canale_data` |
| `--utc` | date in UTC invece dell'ora locale |
| `--api-key KEY` | chiave passata a riga di comando |

Formati di link accettati: `youtu.be/ID`, `watch?v=ID`, `/shorts/ID`, `/live/ID`,
`/embed/ID`, con qualsiasi parametro extra (`si=`, `t=`, `list=`), oppure l'ID nudo.

### Codici di uscita

| Codice | Significato |
|---|---|
| 0 | completato |
| 1 | errore generico |
| 2 | URL o video non valido / non accessibile |
| 3 | API key mancante o rifiutata |
| 4 | quota esaurita prima di iniziare |
| 5 | report **parziale** (quota esaurita, rete o interruzione manuale a metà lavoro) |

## Struttura del report

```
 1  REPORT COMMENTI YOUTUBE                      <- banda blu scuro
 3  Link video      | https://...                <- etichette blu, valori azzurri
 4  Canale          | ...
 5  Titolo video    | ...
 6  Data uscita     | 14/03/2024 18:30
 7  Visualizzazioni | 1.234.567
 8  Commenti dichiarati da YouTube | 4.812
 9  Commenti estratti              | 4.790 (4.102 commenti + 688 risposte)
10  Estrazione del                 | 23/09/2026 11:45
12  COMMENTI
13  Utente | Commento | Data commento | Like | Tipo | Risposta a
14  mario  | ...      | 12/04/2024 09:14 | 12 | Commento |
15    luca | ...      | 12/04/2024 10:02 |  3 | Risposta | mario
```

Le righe delle risposte sono rientrate e su sfondo ambra, i commenti su sfondo bianco/azzurro
alternato. Il foglio scorre per intero, metadati compresi (nessun blocco dei riquadri).
L'intestazione dei commenti ha i filtri automatici attivi; date e numeri sono valori reali,
quindi ordinabili e filtrabili.

**"Commenti dichiarati" ≠ "Commenti estratti" è normale**: YouTube conteggia anche commenti
rimossi, in attesa di revisione o classificati come spam, che l'API non restituisce.

## Comportamenti particolari

- **Commenti disabilitati**: il report viene generato comunque, con i metadati e una nota.
- **Interruzione (Ctrl+C), quota esaurita o errore di rete**: viene salvato ciò che è già
  stato scaricato, con una nota in evidenza nel report (exit code 5).
- **Errori temporanei dell'API** (500/503/rate limit): fino a 5 tentativi con attesa crescente.
- **File già aperto in Excel**: viene salvato con suffisso `_2` invece di fallire.
- **Commenti che iniziano con `=`**: scritti come testo, non interpretati come formule.
- **Video enormi** (oltre 50.000 commenti): scrittura in streaming per contenere la memoria.
- **Commenti oltre 32.767 caratteri**: troncati (limite di Excel) con marcatore `…[troncato]`.

## Test

```bash
python -m unittest discover -s tests
```

Tutti i test sono offline, non richiedono API key né connessione.

## Nota sui dati personali

Il report contiene nomi utente e testi pubblicati da persone reali: conservalo e condividilo
solo per lo scopo per cui lo hai estratto.
