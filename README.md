# 🍷 Landing Menu Degustazione — Antica Trattoria Cirio

> **Cos'è questa landing.** Campagna dedicata al *Menu Degustazione da Cirio*
> (45 € a persona, valido fino a fine agosto). È nata come copia della landing
> prenotazioni Cirio, con in più due blocchi:
>
> - **Blocco offerta** sotto l'hero → si modifica da `CONFIG.degustazione`
> - **Galleria di 3 foto** sotto il form → si modifica da `CONFIG.galleria`
>
> Le foto vanno messe in questa cartella come `foto-1.jpg`, `foto-2.jpg`,
> `foto-3.jpg`. Finché mancano, al loro posto compare un segnaposto discreto
> "Foto in arrivo": la pagina non si rompe mai.
>
> Le prenotazioni finiscono nello stesso Google Sheet della landing Cirio, ma
> sono riconoscibili dal campo *origine* = `Landing Degustazione`.

---

> ## 👉 Come usare questo template
> Questo repository è un **modello riutilizzabile**. Per creare la landing di un
> nuovo cliente **non modificare questo repo**: usa il pulsante verde
> **"Use this template" → "Create a new repository"** in alto.
> Otterrai una copia indipendente; lì apri `index.html`, modifica solo il blocco
> **`CONFIG`** con i dati del cliente e il sito si pubblica da solo su GitHub Pages.
> Dettagli nella sezione **§1** qui sotto.

Template **single-file** (`index.html`) di landing page per ristoranti, il cui unico
scopo è **far prenotare un tavolo**. Nessuna dipendenza, nessun build: si apre con un
doppio click. Per ogni nuovo cliente si modifica **solo l'oggetto `CONFIG`** in cima al
file `index.html`.

Sezioni della pagina: **Hero** (logo + titolo + sottotitolo + CTA) → **Form di
prenotazione intelligente** (con gestione capienza reale) → **Footer** (contatti, orari,
mappa, social).

---

## 1. Creare una landing per un nuovo cliente

Apri `index.html`, trova il blocco `const CONFIG = { … }` (è tutto commentato in
italiano) e modifica solo questi campi:

| Blocco `CONFIG` | Cosa contiene |
|---|---|
| `brand` | Nome ristorante, tagline, `logo` (path immagine) o `iniziali` (fallback). |
| `colori` | Palette. Default nero/bianco/oro FML. `ctaBg` / `ctaText` = colore del bottone. |
| `hero` | `headline` (usa `<em>parola</em>` per l'oro corsivo), `sottotitolo`, testo CTA. |
| `contatti` | Indirizzo, telefono, email, link mappe, social, righe orari del footer. |
| `prenotazioni` | **Capienza, fasce orarie, giorni di chiusura** (vedi §2). |
| `integrazione` | URL del Google Sheet del cliente (vedi §3). |
| `testi` | Etichette fisse dell'interfaccia (di norma non serve toccarle). |

> Il logo: metti il file immagine nella stessa cartella e imposta `brand.logo: "logo.png"`.
> Se lo lasci vuoto (`""`), viene mostrato il cerchio con le iniziali di `brand.iniziali`.

---

## 2. Capienza e fasce orarie

Tutto dentro `CONFIG.prenotazioni`:

```js
capienzaCoperti: 50,   // coperti totali disponibili PER OGNI fascia oraria
fasceOrarie: ["19:00","19:30","20:00","20:30","21:00","21:30","22:00"],
giorniChiusura: [1],   // 0=Dom 1=Lun … 6=Sab  → qui: chiuso il lunedì ([] = mai)
maxPersonePrenotazione: 20,
sogliaUltimiPosti: 8,  // sotto N coperti liberi → badge "Ultimi posti"
giorniPrenotabili: 60, // finestra di prenotazione (giorni da oggi)
```

### Come funziona la disponibilità (il "grigetto")
- Ogni fascia oraria ha un tetto pari a `capienzaCoperti`.
- Il sistema somma i coperti già prenotati per **data + fascia**.
- Se una fascia è piena → diventa **grigia, non cliccabile**, badge **"Completo"**.
- Considera anche le **persone della prenotazione in corso**: se restano 3 posti e
  l'utente ha scelto 4 persone, quella fascia diventa grigia ("Solo 3 posti"). La
  disponibilità viene **ricalcolata ogni volta che si cambia il numero di persone**.
- Quando restano pochi posti (≤ `sogliaUltimiPosti`) → badge **"Ultimi posti"**.

### Dati di esempio per la demo
`prenotazioni.occupatiDemo` contiene coperti finti già occupati, così vedi subito la
logica funzionare. Le chiavi `"GIORNO+1"`, `"GIORNO+2"` vengono convertite in date reali
(oggi + N) grazie a `ancoraDemo: true`. **In produzione svuota `occupatiDemo`** (`{}`):
i conteggi arriveranno dallo Sheet / dal DB.

Le prenotazioni fatte dall'utente vengono salvate in **localStorage**, così ricaricando la
pagina i conteggi restano coerenti e le fasce piene restano grigie.

---

## 3. Collegare il Google Sheet (uno diverso per ogni ristorante)

Ogni landing invia le prenotazioni a uno **Sheet dedicato** tramite una **Web App di
Google Apps Script**. Serve solo incollare un URL in `CONFIG.integrazione.googleSheetsWebAppUrl`.

### Passi (una volta per cliente)
1. Crea un nuovo Google Sheet per il ristorante. Prima riga (intestazioni):
   `Data | Ora | Persone | Nome | Telefono | Email | Richieste | Privacy | Creata`
2. Menu **Estensioni → Apps Script** e incolla:

   ```js
   function doPost(e) {
     const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
     const d = JSON.parse(e.postData.contents);
     sheet.appendRow([d.data, d.ora, d.persone, d.nome, d.telefono, d.email, d.richieste, d.privacy, d.creata]);
     return ContentService.createTextOutput(JSON.stringify({ok:true}))
                          .setMimeType(ContentService.MimeType.JSON);
   }

   // (Opzionale) legge i coperti occupati per far funzionare il grigetto "dal vivo".
   // Attiva anche CONFIG.integrazione.leggiOccupatiDaSheets = true.
   function doGet(e) {
     const rows = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getDataRange().getValues();
     const out = {};
     for (let i = 1; i < rows.length; i++) {
       const [data, ora, persone] = rows[i];
       if (!data) continue;
       const k = Utilities.formatDate(new Date(data), 'GMT', 'yyyy-MM-dd');
       out[k] = out[k] || {};
       out[k][ora] = (out[k][ora] || 0) + Number(persone || 0);
     }
     return ContentService.createTextOutput(JSON.stringify(out))
                          .setMimeType(ContentService.MimeType.JSON);
   }
   ```
3. **Distribuisci → Nuova distribuzione → Tipo: App web**, accesso *"Chiunque"*.
   Copia l'URL `…/exec` e incollalo in:
   ```js
   integrazione: { googleSheetsWebAppUrl: "https://script.google.com/macros/s/XXXX/exec" }
   ```
4. (Opzionale) per leggere le disponibilità reali dallo Sheet all'avvio:
   `integrazione.leggiOccupatiDaSheets: true`.

Se lasci `googleSheetsWebAppUrl: ""`, la landing funziona in **solo-locale** (demo): salva
in localStorage e mostra la conferma, senza invii esterni.

---

## 4. Passare a un DB reale (Supabase) in futuro

La logica di disponibilità è **isolata in funzioni pure e documentate** dentro
`index.html`. Per migrare basta cambiare una funzione, senza toccare la UI:

| Funzione | Ruolo | Cosa cambiare per Supabase |
|---|---|---|
| `getPostiOccupati(data, ora)` | Coperti già occupati per data+fascia | Sostituire con `SELECT COALESCE(SUM(coperti),0) FROM bookings WHERE data=? AND ora=?`. |
| `getDisponibilita(data, ora)` | `capienza − occupati` | Nessuna modifica. |
| `isFasciaDisponibile(data, ora, persone)` | Fascia selezionabile? | Nessuna modifica. |
| `salvaPrenotazione(prenotazione)` | Persiste la prenotazione | Sostituire l'invio con un `INSERT INTO bookings (…)`. |

Nel codice trovi i commenti che segnano **esattamente** i punti da sostituire (cerca
`IN PRODUZIONE` / `QUANDO PASSERAI A SUPABASE`). La tabella suggerita:

```sql
create table bookings (
  id bigint generated always as identity primary key,
  data date not null,
  ora  text not null,        -- "HH:MM"
  coperti int not null,
  nome text, telefono text, richieste text,
  creata timestamptz default now()
);
```

---

## Requisiti coperti
- ✅ Single file HTML autonomo, zero build, apribile con doppio click.
- ✅ Responsive mobile-first, accessibile (label, focus states, aria).
- ✅ Validazione: data non passata, giorni di chiusura e date oltre finestra non
  selezionabili, telefono valido, persone > 0.
- ✅ Disponibilità per capienza + persone, con stati "Completo" / "Ultimi posti".
- ✅ Persistenza localStorage + gancio Google Sheets per-cliente.
