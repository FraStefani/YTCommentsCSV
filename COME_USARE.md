# Come generare il file Excel dei commenti — guida passo a passo

Guida pratica: dal link di un video YouTube al file Excel colorato, in 3 passaggi.
Non serve sapere programmare: si tratta di incollare un comando.

---

## Prima volta soltanto: controlla che tutto sia pronto

1. Apri **PowerShell** (tasto Windows, scrivi `powershell`, Invio) e incolla:
   ```powershell
   python --version
   ```
   Deve rispondere `Python 3.14.5` (o una versione superiore a 3.10).

2. Controlla che esista il file della chiave API. Incolla:
   ```powershell
   Test-Path "C:\workspace\PP\YTCommentsCSV\.env"
   ```
   Deve rispondere `True`. Se rispondesse `False`, la chiave è stata cancellata:
   vedi la sezione *Se manca la chiave API* in fondo.

Fatti questi due controlli una volta, non servono più.

---

## I 3 passaggi da ripetere ogni volta

### 1. Copia il link del video

Su YouTube: pulsante **Condividi** → **Copia**.
Va bene qualsiasi formato di link, non devi ripulirlo:

- `https://youtu.be/ByZrTRph2-4?si=SuKq_vqCI-q9MRt7`
- `https://www.youtube.com/watch?v=ByZrTRph2-4&t=42s`
- `https://www.youtube.com/shorts/ByZrTRph2-4`
- anche solo `ByZrTRph2-4`

### 2. Apri PowerShell nella cartella del programma

Incolla questo comando (ti porta nella cartella giusta):

```powershell
cd "C:\workspace\PP\YTCommentsCSV"
```

> Scorciatoia: apri la cartella in Esplora file, clicca sulla barra dell'indirizzo
> in alto, scrivi `powershell` e premi Invio. PowerShell si apre già nella cartella.

### 3. Lancia il comando con il tuo link

Scrivi questo, **sostituendo il link tra le virgolette** con quello che hai copiato:

```powershell
python ytcomments.py "INCOLLA-QUI-IL-LINK" --out "Excel generati"
```

Esempio reale, già pronto:

```powershell
python ytcomments.py "https://youtu.be/ByZrTRph2-4?si=SuKq_vqCI-q9MRt7" --out "Excel generati"
```

**Le virgolette servono sempre**: i link contengono caratteri (`&`, `?`) che senza
virgolette PowerShell interpreta come comandi e il programma non partirebbe.

---

## Cosa vedi mentre lavora

```
  scaricati 40 commenti...
Canale ......... InnTale
Video .......... REAMI DIMENTICATI: Gabbo & Curren
Visualizzazioni  14.852
Commenti ....... 40 (dichiarati da YouTube)

Estratti ....... 40 (31 commenti + 9 risposte)
Quota usata .... 2 unità (chiamate API)
File ........... Excel generati\InnTale_2026-09-23.xlsx
```

L'ultima riga è il percorso del file appena creato: si chiama
`<nome canale>_<data di oggi>.xlsx` e sta nella cartella **Excel generati**.
Aprilo con un doppio clic.

Se generi due volte lo stesso canale nello stesso giorno, il secondo file prende
automaticamente il suffisso `_2` (`InnTale_2026-09-23_2.xlsx`): niente viene sovrascritto.

---

## Varianti utili

| Se vuoi... | Aggiungi al comando |
|---|---|
| fare una prova veloce su pochi commenti | `--max 50` |
| escludere le risposte ai commenti | `--no-replies` |
| i commenti più rilevanti invece che i più recenti | `--order relevance` |
| salvare in un'altra cartella | `--out "C:\altra\cartella"` |
| dare un nome tuo al file | `--name "report riunione"` |
| date in UTC invece dell'ora italiana | `--utc` |

Se la cartella indicata con `--out` non esiste, viene creata automaticamente.

Si possono combinare. Esempio — prova rapida da 20 commenti senza risposte:

```powershell
python ytcomments.py "https://youtu.be/ByZrTRph2-4" --max 20 --no-replies --out "Excel generati"
```

Per rivedere l'elenco completo delle opzioni:

```powershell
python ytcomments.py --help
```

---

## Se qualcosa va storto

| Messaggio a schermo | Cosa significa e cosa fare |
|---|---|
| `Non sembra un link YouTube: ...` oppure `Impossibile ricavare l'ID video da: ...` | il link è incompleto o non è di YouTube: ricopialo con **Condividi → Copia** |
| `Nessun video con ID ...: potrebbe essere privato, rimosso...` | il video non è pubblico o il link è di un altro video |
| `ERRORE: API key di YouTube non trovata` | manca il file `.env`: vedi la sezione qui sotto |
| `API key rifiutata da Google: ...` | la chiave è stata revocata o limitata: creane una nuova |
| `Quota giornaliera dell'API esaurita` | hai superato il limite gratuito di Google: **il file viene salvato comunque** con i commenti già scaricati e una nota in evidenza. La quota si azzera a mezzanotte ora del Pacifico (9:00 del mattino in Italia) |
| `impossibile scrivere in ...` | hai il file già aperto in Excel: chiudilo, oppure lascia fare — viene salvato con suffisso `_2` |
| `ATTENZIONE, report PARZIALE` | il lavoro si è interrotto a metà (rete, quota, Ctrl+C) ma **quello che era già stato scaricato è nel file**, con la nota scritta dentro il report |

Puoi interrompere in qualsiasi momento con **Ctrl+C**: i commenti già scaricati
vengono salvati comunque, non perdi il lavoro.

### Numeri diversi tra "dichiarati" e "estratti"

Non è un errore: YouTube nel conteggio pubblico include anche commenti rimossi,
in attesa di revisione o segnalati come spam, che l'API non restituisce.
Nel report trovi entrambi i numeri, uno sotto l'altro, proprio per questo.

### Se manca la chiave API

1. Vai su <https://console.cloud.google.com/> e crea un progetto.
2. **API e servizi → Libreria →** cerca *YouTube Data API v3* → **Abilita**.
3. **API e servizi → Credenziali → Crea credenziali → Chiave API** → copia la chiave.
4. Nella cartella del programma, crea un file di testo chiamato `.env` contenente
   una sola riga:
   ```
   YOUTUBE_API_KEY=AIza...la-tua-chiave...
   ```
   In alternativa puoi passarla direttamente nel comando con `--api-key AIza...`.

La chiave è gratuita e il limite (10.000 unità al giorno, circa un milione di
commenti) è molto più alto del normale uso. Non viene mai scritta nel file Excel
né nei messaggi di errore.

---

## Promemoria in una riga

```powershell
cd "C:\workspace\PP\YTCommentsCSV"; python ytcomments.py "LINK" --out "Excel generati"
```
